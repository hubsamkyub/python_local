# 데이터베이스 관계 분석 도구
# Database Relationship Analyzer

import pyodbc
import pandas as pd
import networkx as nx
import plotly.graph_objects as go
import plotly.express as px
from plotly.subplots import make_subplots
from difflib import SequenceMatcher
from collections import defaultdict
import re
import json
from typing import Dict, List, Tuple, Set
import warnings
warnings.filterwarnings('ignore')

class DatabaseConfig:
    """데이터베이스 연결 설정 클래스"""
    
    def __init__(self, server: str, database: str, username: str, password: str):
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        
    def get_connection_string(self) -> str:
        """SQL Server 연결 문자열 생성"""
        return (
            f"DRIVER={{ODBC Driver 17 for SQL Server}};"
            f"SERVER={self.server};"
            f"DATABASE={self.database};"
            f"UID={self.username};"
            f"PWD={self.password};"
            f"TrustServerCertificate=yes;"
        )

class DatabaseRelationAnalyzer:
    """데이터베이스 테이블 관계 분석기"""
    
    def __init__(self, config: DatabaseConfig):
        self.config = config
        self.connection = None
        self.tables_info = {}
        self.relationships = []
        self.graph = nx.DiGraph()
        
    def connect(self) -> bool:
        """데이터베이스 연결"""
        try:
            self.connection = pyodbc.connect(self.config.get_connection_string())
            print("✅ 데이터베이스 연결 성공")
            return True
        except Exception as e:
            print(f"❌ 데이터베이스 연결 실패: {e}")
            return False
    
    def get_all_tables_columns(self) -> pd.DataFrame:
        """모든 테이블의 컬럼 정보 조회"""
        query = """
        SELECT 
            t.name AS table_name,
            c.name AS column_name,
            ty.name AS data_type,
            c.is_nullable,
            c.column_id
        FROM sys.tables t
        INNER JOIN sys.columns c ON t.object_id = c.object_id
        INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
        WHERE OBJECT_SCHEMA_NAME(t.object_id) = 'dbo'
        ORDER BY t.name, c.column_id;
        """
        return pd.read_sql(query, self.connection)
    
    def get_id_columns(self) -> pd.DataFrame:
        """ID로 끝나는 컬럼들만 조회"""
        query = """
        SELECT 
            t.name AS table_name,
            c.name AS column_name,
            ty.name AS data_type
        FROM sys.tables t
        INNER JOIN sys.columns c ON t.object_id = c.object_id
        INNER JOIN sys.types ty ON c.user_type_id = ty.user_type_id
        WHERE OBJECT_SCHEMA_NAME(t.object_id) = 'dbo'
            AND (c.name LIKE '%ID' OR c.name LIKE '%Id' OR c.name LIKE '%_id')
        ORDER BY t.name, c.column_id;
        """
        return pd.read_sql(query, self.connection)
    
    def analyze_column_patterns(self) -> Dict[str, List[str]]:
        """컬럼명 패턴 분석"""
        id_columns = self.get_id_columns()
        
        # 컬럼명별로 사용하는 테이블들 그룹핑
        column_groups = defaultdict(list)
        for _, row in id_columns.iterrows():
            column_groups[row['column_name']].append(row['table_name'])
        
        # 2개 이상 테이블에서 사용되는 컬럼만 반환
        return {col: tables for col, tables in column_groups.items() if len(tables) > 1}
    
    def find_similar_columns(self, threshold: float = 0.6) -> List[Tuple[str, str, float]]:
        """유사한 컬럼명들 찾기 (예: RewardGroupID ↔ BonusRewardGroupID)"""
        id_columns = self.get_id_columns()
        unique_columns = id_columns['column_name'].unique()
        
        similar_pairs = []
        for i, col1 in enumerate(unique_columns):
            for col2 in unique_columns[i+1:]:
                similarity = SequenceMatcher(None, col1.lower(), col2.lower()).ratio()
                if similarity >= threshold:
                    similar_pairs.append((col1, col2, similarity))
        
        return sorted(similar_pairs, key=lambda x: x[2], reverse=True)
    
    def verify_data_relationship(self, table1: str, col1: str, table2: str, col2: str) -> Dict:
        """실제 데이터로 관계 검증"""
        try:
            query = f"""
            SELECT 
                'table1_only' as status,
                COUNT(DISTINCT t1.{col1}) as count
            FROM [{self.config.database}].[dbo].{table1} t1
            WHERE NOT EXISTS (
                SELECT 1 FROM [{self.config.database}].[dbo].{table2} t2 
                WHERE t2.{col2} = t1.{col1}
            )
            
            UNION ALL
            
            SELECT 
                'table2_only' as status,
                COUNT(DISTINCT t2.{col2}) as count
            FROM [{self.config.database}].[dbo].{table2} t2
            WHERE NOT EXISTS (
                SELECT 1 FROM [{self.config.database}].[dbo].{table1} t1 
                WHERE t1.{col1} = t2.{col2}
            )
            
            UNION ALL
            
            SELECT 
                'both_exist' as status,
                COUNT(DISTINCT t1.{col1}) as count
            FROM [{self.config.database}].[dbo].{table1} t1
            WHERE EXISTS (
                SELECT 1 FROM [{self.config.database}].[dbo].{table2} t2 
                WHERE t2.{col2} = t1.{col1}
            );
            """
            
            result = pd.read_sql(query, self.connection)
            
            # 결과를 딕셔너리로 변환
            result_dict = {row['status']: row['count'] for _, row in result.iterrows()}
            
            # 관계 강도 계산
            table1_only = result_dict.get('table1_only', 0)
            table2_only = result_dict.get('table2_only', 0)
            both_exist = result_dict.get('both_exist', 0)
            
            total_table1 = table1_only + both_exist
            total_table2 = table2_only + both_exist
            
            if total_table1 == 0 and total_table2 == 0:
                relationship_type = "no_data"
            elif table1_only == 0 and table2_only == 0 and both_exist > 0:
                relationship_type = "perfect_match"
            elif table1_only == 0 and both_exist > 0:
                relationship_type = "table1_subset_of_table2"
            elif table2_only == 0 and both_exist > 0:
                relationship_type = "table2_subset_of_table1"
            else:
                relationship_type = "partial_match"
            
            return {
                'table1': table1,
                'column1': col1,
                'table2': table2,
                'column2': col2,
                'table1_only': table1_only,
                'table2_only': table2_only,
                'both_exist': both_exist,
                'total_table1': total_table1,
                'total_table2': total_table2,
                'relationship_type': relationship_type,
                'match_ratio': both_exist / max(total_table1, total_table2) if max(total_table1, total_table2) > 0 else 0
            }
            
        except Exception as e:
            print(f"❌ 관계 검증 실패 ({table1}.{col1} ↔ {table2}.{col2}): {e}")
            return None
    
    def analyze_all_relationships(self) -> List[Dict]:
        """모든 관계 분석"""
        print("📊 컬럼 패턴 분석 중...")
        column_patterns = self.analyze_column_patterns()
        
        print("🔍 유사 컬럼명 분석 중...")
        similar_columns = self.find_similar_columns()
        
        relationships = []
        
        # 1. 같은 컬럼명을 사용하는 테이블들 간의 관계 분석
        print("🔗 같은 컬럼명 관계 분석 중...")
        for column_name, tables in column_patterns.items():
            for i, table1 in enumerate(tables):
                for table2 in tables[i+1:]:
                    relationship = self.verify_data_relationship(table1, column_name, table2, column_name)
                    if relationship and relationship['both_exist'] > 0:
                        relationships.append(relationship)
        
        # 2. 유사한 컬럼명들 간의 관계 분석
        print("🔗 유사 컬럼명 관계 분석 중...")
        id_columns_df = self.get_id_columns()
        
        for col1, col2, similarity in similar_columns[:10]:  # 상위 10개만 분석
            # col1을 사용하는 테이블들
            tables1 = id_columns_df[id_columns_df['column_name'] == col1]['table_name'].tolist()
            # col2를 사용하는 테이블들
            tables2 = id_columns_df[id_columns_df['column_name'] == col2]['table_name'].tolist()
            
            for table1 in tables1:
                for table2 in tables2:
                    if table1 != table2:  # 같은 테이블이 아닌 경우
                        relationship = self.verify_data_relationship(table1, col1, table2, col2)
                        if relationship and relationship['both_exist'] > 0:
                            relationship['similarity_score'] = similarity
                            relationships.append(relationship)
        
        self.relationships = relationships
        return relationships
    
    def build_relationship_graph(self) -> nx.DiGraph:
        """관계 그래프 구축"""
        self.graph = nx.DiGraph()
        
        for rel in self.relationships:
            if rel['both_exist'] > 0:
                # 엣지 가중치는 매칭되는 데이터 수
                weight = rel['both_exist']
                edge_type = rel['relationship_type']
                
                self.graph.add_edge(
                    rel['table1'], 
                    rel['table2'], 
                    weight=weight,
                    relationship_type=edge_type,
                    column1=rel['column1'],
                    column2=rel['column2'],
                    match_ratio=rel['match_ratio']
                )
        
        return self.graph
    
    def visualize_relationships(self, min_weight: int = 1) -> go.Figure:
        """관계도 시각화"""
        if not self.graph:
            self.build_relationship_graph()
        
        # 최소 가중치 이상의 엣지만 필터링
        filtered_edges = [(u, v, d) for u, v, d in self.graph.edges(data=True) if d['weight'] >= min_weight]
        
        if not filtered_edges:
            print("표시할 관계가 없습니다.")
            return None
        
        # 노드 위치 계산
        pos = nx.spring_layout(self.graph, k=3, iterations=50)
        
        # 엣지 데이터 준비
        edge_x = []
        edge_y = []
        edge_info = []
        
        for edge in filtered_edges:
            x0, y0 = pos[edge[0]]
            x1, y1 = pos[edge[1]]
            edge_x.extend([x0, x1, None])
            edge_y.extend([y0, y1, None])
            
            # 엣지 정보
            edge_info.append(f"{edge[0]}.{edge[2]['column1']} ↔ {edge[1]}.{edge[2]['column2']}<br>"
                           f"매칭 데이터: {edge[2]['weight']}개<br>"
                           f"관계 타입: {edge[2]['relationship_type']}<br>"
                           f"매칭 비율: {edge[2]['match_ratio']:.2%}")
        
        # 노드 데이터 준비
        node_x = []
        node_y = []
        node_text = []
        node_sizes = []
        
        for node in self.graph.nodes():
            x, y = pos[node]
            node_x.append(x)
            node_y.append(y)
            node_text.append(node)
            
            # 노드 크기는 연결된 엣지 수에 비례
            degree = self.graph.degree(node)
            node_sizes.append(max(20, degree * 5))
        
        # 플롯 생성
        fig = go.Figure()
        
        # 엣지 추가
        fig.add_trace(go.Scatter(
            x=edge_x, y=edge_y,
            line=dict(width=2, color='lightblue'),
            hoverinfo='none',
            mode='lines'
        ))
        
        # 노드 추가
        fig.add_trace(go.Scatter(
            x=node_x, y=node_y,
            mode='markers+text',
            hoverinfo='text',
            text=node_text,
            textposition="middle center",
            hovertext=node_text,
            marker=dict(
                size=node_sizes,
                color='lightcoral',
                line=dict(width=2, color='white')
            )
        ))
        
        fig.update_layout(
            title={'text': '데이터베이스 테이블 관계도', 'font': {'size': 16}},
            showlegend=False,
            hovermode='closest',
            margin=dict(b=20,l=5,r=5,t=40),
            annotations=[ dict(
                text="테이블을 클릭하면 상세 정보를 볼 수 있습니다",
                showarrow=False,
                xref="paper", yref="paper",
                x=0.005, y=-0.002,
                xanchor='left', yanchor='bottom',
                font=dict(color='gray', size=12)
            )],
            xaxis=dict(showgrid=False, zeroline=False, showticklabels=False),
            yaxis=dict(showgrid=False, zeroline=False, showticklabels=False)
        )
        
        return fig
    
    def find_impact_analysis(self, table_name: str, column_name: str, value: any) -> Dict:
        """특정 데이터 삭제/추가 시 영향 분석"""
        impact = {
            'deletion_impact': [],
            'required_additions': [],
            'data_integrity_issues': []
        }
        
        # 해당 테이블과 관련된 관계들 찾기
        related_relationships = [
            rel for rel in self.relationships 
            if (rel['table1'] == table_name and rel['column1'] == column_name) or
               (rel['table2'] == table_name and rel['column2'] == column_name)
        ]
        
        for rel in related_relationships:
            if rel['table1'] == table_name and rel['column1'] == column_name:
                # 다른 테이블에서 이 값을 참조하는지 확인
                impact['deletion_impact'].append({
                    'affected_table': rel['table2'],
                    'affected_column': rel['column2'],
                    'relationship_type': rel['relationship_type']
                })
            elif rel['table2'] == table_name and rel['column2'] == column_name:
                # 이 값이 다른 테이블에 존재해야 하는지 확인
                impact['required_additions'].append({
                    'required_table': rel['table1'],
                    'required_column': rel['column1'],
                    'relationship_type': rel['relationship_type']
                })
        
        return impact
    
    def generate_report(self) -> str:
        """분석 결과 리포트 생성"""
        if not self.relationships:
            return "분석 결과가 없습니다. analyze_all_relationships()를 먼저 실행해주세요."
        
        report = "# 데이터베이스 관계 분석 리포트\n\n"
        
        # 전체 통계
        total_relationships = len(self.relationships)
        perfect_matches = len([r for r in self.relationships if r['relationship_type'] == 'perfect_match'])
        partial_matches = len([r for r in self.relationships if r['relationship_type'] == 'partial_match'])
        
        report += f"## 📊 전체 통계\n"
        report += f"- 발견된 관계: {total_relationships}개\n"
        report += f"- 완벽한 매칭: {perfect_matches}개\n"
        report += f"- 부분적 매칭: {partial_matches}개\n\n"
        
        # 주요 관계들
        report += f"## 🔗 주요 관계 목록\n"
        for rel in sorted(self.relationships, key=lambda x: x['both_exist'], reverse=True)[:20]:
            report += f"- **{rel['table1']}.{rel['column1']}** ↔ **{rel['table2']}.{rel['column2']}**\n"
            report += f"  - 매칭 데이터: {rel['both_exist']}개\n"
            report += f"  - 관계 타입: {rel['relationship_type']}\n"
            report += f"  - 매칭 비율: {rel['match_ratio']:.2%}\n\n"
        
        # 데이터 정합성 문제
        integrity_issues = [r for r in self.relationships if r['table1_only'] > 0 or r['table2_only'] > 0]
        if integrity_issues:
            report += f"## ⚠️ 데이터 정합성 문제\n"
            for issue in integrity_issues[:10]:
                if issue['table1_only'] > 0:
                    report += f"- **{issue['table1']}.{issue['column1']}**: {issue['table1_only']}개 값이 {issue['table2']}에 없음\n"
                if issue['table2_only'] > 0:
                    report += f"- **{issue['table2']}.{issue['column2']}**: {issue['table2_only']}개 값이 {issue['table1']}에 없음\n"
            report += "\n"
        
        return report
    
    def close(self):
        """데이터베이스 연결 종료"""
        if self.connection:
            self.connection.close()
            print("✅ 데이터베이스 연결 종료")

# 사용 예시 함수들
def create_config():
    """설정 생성 도우미 함수"""
    print("=== 데이터베이스 연결 설정 ===")
    server = input("서버 주소: ")
    database = input("데이터베이스명 (기본값: GameDataDB_limdongbin): ") or "GameDataDB_limdongbin"
    username = input("사용자명: ")
    password = input("비밀번호: ")
    
    return DatabaseConfig(server, database, username, password)

def main():
    """메인 실행 함수"""
    # 1. 설정 생성
    # config = create_config()  # 실제 사용 시 주석 해제
    
    # 예시용 더미 설정 (실제 사용 시 위 라인으로 교체)
    config = DatabaseConfig(
        server="192.168.0.187,24336",
        database="GameDataDB_limdongbin", 
        username="ProjectH",
        password="Projecth123#"
    )
    
    # 2. 분석기 생성 및 연결
    analyzer = DatabaseRelationAnalyzer(config)
    
    if not analyzer.connect():
        return
    
    try:
        # 3. 관계 분석 실행
        print("🔍 관계 분석 시작...")
        relationships = analyzer.analyze_all_relationships()
        
        print(f"✅ {len(relationships)}개 관계 발견")
        
        # 4. 시각화
        print("📈 관계도 생성 중...")
        fig = analyzer.visualize_relationships(min_weight=10)  # 10개 이상 매칭되는 관계만 표시
        if fig:
            fig.show()
        
        # 5. 리포트 생성
        print("📄 리포트 생성 중...")
        report = analyzer.generate_report()
        print(report)
        
    finally:
        analyzer.close()

if __name__ == "__main__":
    main()