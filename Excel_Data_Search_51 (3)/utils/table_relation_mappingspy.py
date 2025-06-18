# utils/enum_mappings.py

import os
import json
import logging
from collections import defaultdict

# 로거 설정
logger = logging.getLogger(__name__)

# ConditionType 매핑 (enum_typecode 관계에서 추출)
CONDITION_TYPE_MAPPINGS = {
    "BattleRage": {"table": "BattleRage", "column": "Id", "description": "배틀 레이지"},
    "CalendarTemplate": {"table": "CalendarTemplate", "column": "Id", "description": "캘린더 템플릿"},
    "ConditionTemplate": {"table": "ConditionTemplate", "column": "Id", "description": "조건 템플릿"},
    "ContentsExterminate": {"table": "ContentsExterminate", "column": "Id", "description": "컨텐츠 퇴치"},
}

# RewardType 매핑 (enum_typecode 관계에서 추출)
REWARD_TYPE_MAPPINGS = {
    "BoxTemplate": {"table": "BoxTemplate", "column": "Id", "description": "박스 템플릿"},
    "CalendarRewardTemplate": {"table": "CalendarRewardTemplate", "column": "Id", "description": "캘린더 보상 템플릿"},
    "CashShopTemplate": {"table": "CashShopTemplate", "column": "Id", "description": "캐시샵 템플릿"},
    "CollectionAchieveReward": {"table": "CollectionAchieveReward", "column": "Id", "description": "컬렉션 달성 보상"},
    "EventHotTimeReward": {"table": "EventHotTimeReward", "column": "Id", "description": "이벤트 핫타임 보상"},
    "GachaBoxTemplate": {"table": "GachaBoxTemplate", "column": "Id", "description": "가챠 박스 템플릿"},
    "GachaStepRewardTemplate": {"table": "GachaStepRewardTemplate", "column": "Id", "description": "가챠 단계 보상 템플릿"},
    "HallofFameTemplate": {"table": "HallofFameTemplate", "column": "Id", "description": "명예의 전당 템플릿"},
    "ItemDropInfo": {"table": "ItemDropInfo", "column": "Id", "description": "아이템 드롭 정보"},
    "MinigameRunnerScore": {"table": "MinigameRunnerScore", "column": "Id", "description": "미니게임 러너 점수"},
    "NPCShopTemplate": {"table": "NPCShopTemplate", "column": "Id", "description": "NPC 상점 템플릿"},
    "PointRewardTemplate": {"table": "PointRewardTemplate", "column": "Id", "description": "포인트 보상 템플릿"},
    "RandomShopReward": {"table": "RandomShopReward", "column": "Id", "description": "랜덤 상점 보상"},
    "StageMissionTemplate": {"table": "StageMissionTemplate", "column": "Id", "description": "스테이지 미션 템플릿"},
    "Tutorial": {"table": "Tutorial", "column": "Id", "description": "튜토리얼"},
}

# CostType 매핑 (enum_typecode 관계에서 추출)
COST_TYPE_MAPPINGS = {
    "CalendarTemplate": {"table": "CalendarTemplate", "column": "Id", "description": "캘린더 템플릿"},
    "CashShopTemplate": {"table": "CashShopTemplate", "column": "Id", "description": "캐시샵 템플릿"},
    "ContentsExterminate": {"table": "ContentsExterminate", "column": "Id", "description": "컨텐츠 퇴치"},
    "GachaTemplate": {"table": "GachaTemplate", "column": "Id", "description": "가챠 템플릿"},
    "HeroStorySetTemplate": {"table": "HeroStorySetTemplate", "column": "Id", "description": "영웅 스토리 세트 템플릿"},
    "ItemWorkmanshipTemplate": {"table": "ItemWorkmanshipTemplate", "column": "Id", "description": "아이템 제작 템플릿"},
    "MakeCostGroup": {"table": "MakeCostGroup", "column": "Id", "description": "제작 비용 그룹"},
    "NPCShopTemplate": {"table": "NPCShopTemplate", "column": "Id", "description": "NPC 상점 템플릿"},
    "RandomShopReward": {"table": "RandomShopReward", "column": "Id", "description": "랜덤 상점 보상"},
    "StageTemplate": {"table": "StageTemplate", "column": "Id", "description": "스테이지 템플릿"},
    "SynergyClockworkEnchantGroup": {"table": "SynergyClockworkEnchantGroup", "column": "Id", "description": "시너지 시계장치 인챈트 그룹"},
}

# RewardGroupType 매핑 (enum_typecode 관계에서 추출)
REWARD_GROUP_TYPE_MAPPINGS = {
    "RewardGroupTemplate": {"table": "RewardGroupTemplate", "column": "Id", "description": "보상 그룹 템플릿"},
}

# QuestType 매핑 (enum_typecode 관계에서 추출)
QUEST_TYPE_MAPPINGS = {
    "QuestTemplate": {"table": "QuestTemplate", "column": "Id", "description": "퀘스트 템플릿"},
}

# StageType 매핑 (enum_typecode 관계에서 추출)
STAGE_TYPE_MAPPINGS = {
    "StageScheduleTemplate": {"table": "StageScheduleTemplate", "column": "Id", "description": "스테이지 스케줄 템플릿"},
    "StageTemplate": {"table": "StageTemplate", "column": "Id", "description": "스테이지 템플릿"},
}

# 전체 타입 매핑 딕셔너리
TYPE_MAPPINGS = {
    "ConditionType": CONDITION_TYPE_MAPPINGS,
    "RewardType": REWARD_TYPE_MAPPINGS,
    "CostType": COST_TYPE_MAPPINGS,
    "RewardGroupType": REWARD_GROUP_TYPE_MAPPINGS,
    "QuestType": QUEST_TYPE_MAPPINGS,
    "StageType": STAGE_TYPE_MAPPINGS,
}

# 특정 타입별 상세 매핑 (값별 분류)
REWARD_TYPE_VALUES = {
    # 영웅 관련
    "10": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅"},
    "11": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅 조각"},
    
    # 아이템 관련
    "20": {"table": "ItemTemplate", "column": "TemplateID", "description": "아이템"},
    "21": {"table": "ItemTemplate", "column": "TemplateID", "description": "장비"},
    
    # 재화 관련
    "30": {"table": "GoodsMaxValue", "column": "GoodsType", "description": "재화"},
    
    # 박스 관련
    "40": {"table": "BoxTemplate", "column": "ItemTID", "description": "상자"},
    "41": {"table": "BoxTemplate", "column": "ItemTID", "description": "선택 상자"},
    
    # 티켓 관련
    "50": {"table": "TicketMaxValue", "column": "TicketType", "description": "티켓"},
    
    # 기타
    "70": {"table": "WisdomBookTemplate", "column": "TemplateID", "description": "도감"},
    "80": {"table": "CostumeTemplate", "column": "TemplateID", "description": "코스튬"}
}

COST_TYPE_VALUES = {
    # 영웅 관련
    "10": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅"},
    
    # 아이템 관련
    "20": {"table": "ItemTemplate", "column": "TemplateID", "description": "아이템"},
    
    # 재화 관련
    "30": {"table": "GoodsMaxValue", "column": "GoodsType", "description": "재화"},
    
    # 티켓 관련
    "50": {"table": "TicketMaxValue", "column": "TicketType", "description": "티켓"},
}

CONDITION_TYPE_VALUES = {
    # 스테이지 관련
    "10": {"table": "StageTemplate", "column": "UniqueID", "description": "스테이지 클리어"},
    "11": {"table": "StageTemplate", "column": "StageGroup", "description": "스테이지 그룹 클리어"},
    
    # 퀘스트 관련
    "20": {"table": "QuestTemplate", "column": "TemplateID", "description": "퀘스트 완료"},
    
    # 영웅 관련
    "30": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅 획득"},
    "31": {"table": "HeroTemplate", "column": "BaseHeroID", "description": "영웅 레벨업"},
    
    # 아이템 관련
    "40": {"table": "ItemTemplate", "column": "TemplateID", "description": "아이템 획득"},
    "41": {"table": "ItemTemplate", "column": "TemplateID", "description": "아이템 강화"},
    
    # 기타
    "100": {"table": "Tutorial", "column": "TemplateID", "description": "튜토리얼 완료"},
}

def load_relationships_from_file(file_path):
    """
    JSON 파일에서 테이블 관계 정보를 로드합니다.
    
    Args:
        file_path (str): JSON 파일 경로
        
    Returns:
        dict: 테이블 관계 정보
    """
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            relationships = json.load(f)
            logger.debug(f"관계 정보 로드 성공: {len(relationships)} 테이블")
            return relationships
    except Exception as e:
        logger.error(f"관계 정보 로드 오류: {e}")
        return {}

def extract_enum_mappings(relationships):
    """
    테이블 관계 정보에서 enum_typecode 관계들을 추출하여 매핑 딕셔너리를 생성합니다.
    
    Args:
        relationships (dict): 테이블 관계 정보
        
    Returns:
        dict: 타입별 매핑 정보
    """
    type_mappings = defaultdict(dict)
    
    for source_table, target_tables in relationships.items():
        for target_table, relations in target_tables.items():
            for relation in relations:
                if relation.get("relation_type") == "enum_typecode":
                    filter_value = relation.get("filter_value")
                    if filter_value:
                        # filter_value를 타입명으로 사용
                        type_name = filter_value
                        if type_name not in type_mappings:
                            type_mappings[type_name] = {}
                        
                        type_mappings[type_name][target_table] = {
                            "table": target_table,
                            "column": relation.get("target_column", "Id"),
                            "description": target_table,
                            "source_table": source_table,
                            "source_column": relation.get("source_column", "Value"),
                            "filter_column": relation.get("filter_column")
                        }
    
    return dict(type_mappings)

def get_enum_mapping(type_name, table_name=None):
    """
    특정 enum 타입의 매핑 정보를 반환합니다.
    
    Args:
        type_name (str): enum 타입명 (예: "RewardType", "CostType")
        table_name (str, optional): 특정 테이블의 매핑만 반환
        
    Returns:
        dict: 매핑 정보
    """
    if type_name in TYPE_MAPPINGS:
        mappings = TYPE_MAPPINGS[type_name]
        if table_name and table_name in mappings:
            return mappings[table_name]
        return mappings
    
    logger.warning(f"알 수 없는 enum 타입: {type_name}")
    return {}

def get_enum_value_mapping(type_name, value):
    """
    특정 enum 타입과 값에 해당하는 매핑 정보를 반환합니다.
    
    Args:
        type_name (str): enum 타입명 (예: "RewardType", "CostType")
        value (str|int): enum 값
        
    Returns:
        dict: 매핑 정보 또는 None
    """
    value_str = str(value)
    
    # 값별 매핑에서 검색
    if type_name == "RewardType" and value_str in REWARD_TYPE_VALUES:
        return REWARD_TYPE_VALUES[value_str]
    elif type_name == "CostType" and value_str in COST_TYPE_VALUES:
        return COST_TYPE_VALUES[value_str]
    elif type_name == "ConditionType" and value_str in CONDITION_TYPE_VALUES:
        return CONDITION_TYPE_VALUES[value_str]
    
    logger.warning(f"알 수 없는 {type_name} 값: {value}")
    return None

def get_all_enum_types():
    """
    사용 가능한 모든 enum 타입 목록을 반환합니다.
    
    Returns:
        list: enum 타입명 리스트
    """
    return list(TYPE_MAPPINGS.keys())

def get_tables_for_enum_type(type_name):
    """
    특정 enum 타입에서 사용되는 모든 테이블 목록을 반환합니다.
    
    Args:
        type_name (str): enum 타입명
        
    Returns:
        list: 테이블명 리스트
    """
    if type_name in TYPE_MAPPINGS:
        return list(TYPE_MAPPINGS[type_name].keys())
    return []

def resolve_enum_reference(source_table, filter_column, filter_value, target_value):
    """
    enum 참조를 해결합니다.
    
    Args:
        source_table (str): 소스 테이블명
        filter_column (str): 필터 컬럼명 (예: "AdjustType")
        filter_value (str): 필터 값 (예: "RewardType")
        target_value (str|int): 참조할 값
        
    Returns:
        dict: 참조 해결 결과
    """
    # enum 타입 매핑에서 검색
    mapping = get_enum_value_mapping(filter_value, target_value)
    
    if mapping:
        return {
            "resolved": True,
            "table": mapping["table"],
            "column": mapping["column"],
            "description": mapping["description"],
            "enum_type": filter_value,
            "enum_value": str(target_value)
        }
    
    # 기본 정보 반환
    return {
        "resolved": False,
        "enum_type": filter_value,
        "enum_value": str(target_value),
        "source_table": source_table,
        "filter_column": filter_column
    }

# 특별한 enum 값들 (게임별 커스텀)
CUSTOM_ENUM_VALUES = {
    "RewardType": {
        "1001": {"table": "SpecialReward", "column": "Id", "description": "특별 보상"},
        "1002": {"table": "EventReward", "column": "Id", "description": "이벤트 보상"},
    },
    "CostType": {
        "1001": {"table": "SpecialCost", "column": "Id", "description": "특별 비용"},
    }
}

def generate_enum_mappings_from_json(json_file_path):
    """
    JSON 파일로부터 자동으로 enum 매핑을 생성합니다.
    
    Args:
        json_file_path (str): table_relationships.json 파일 경로
        
    Returns:
        str: Python 코드 문자열
    """
    relationships = load_relationships_from_file(json_file_path)
    if not relationships:
        return ""
    
    enum_mappings = extract_enum_mappings(relationships)
    
    code_lines = [
        "# 자동 생성된 enum 매핑",
        "# Generated from table_relationships.json",
        "",
    ]
    
    for enum_type, mappings in enum_mappings.items():
        mapping_name = f"{enum_type.upper()}_MAPPINGS"
        code_lines.append(f"# {enum_type} 매핑")
        code_lines.append(f"{mapping_name} = {{")
        
        for table_name, mapping_info in mappings.items():
            description = mapping_info.get("description", table_name)
            code_lines.append(f'    "{table_name}": {{')
            code_lines.append(f'        "table": "{mapping_info["table"]}",')
            code_lines.append(f'        "column": "{mapping_info["column"]}",')
            code_lines.append(f'        "description": "{description}"')
            code_lines.append('    },')
        
        code_lines.append("}")
        code_lines.append("")
    
    return "\n".join(code_lines)

def save_enum_mappings_to_file(json_file_path, output_file_path):
    """
    JSON 파일로부터 enum 매핑을 생성하여 Python 파일로 저장합니다.
    
    Args:
        json_file_path (str): 입력 JSON 파일 경로
        output_file_path (str): 출력 Python 파일 경로
    """
    code = generate_enum_mappings_from_json(json_file_path)
    
    if code:
        try:
            with open(output_file_path, 'w', encoding='utf-8') as f:
                f.write(code)
            logger.info(f"enum 매핑 파일 생성 완료: {output_file_path}")
        except Exception as e:
            logger.error(f"파일 저장 오류: {e}")
    else:
        logger.warning("생성할 코드가 없습니다.")

def update_enum_mappings_from_json(json_file_path=None):
    """
    JSON 파일로부터 현재 모듈의 매핑을 업데이트합니다.
    
    Args:
        json_file_path (str, optional): JSON 파일 경로 (없으면 자동 탐색)
    """
    if json_file_path is None:
        # 자동으로 파일 찾기
        search_paths = [
            "table_relationships.json",
            os.path.join(os.getcwd(), "table_relationships.json"),
            os.path.join(os.path.dirname(__file__), "table_relationships.json"),
        ]
        
        for path in search_paths:
            if os.path.exists(path):
                json_file_path = path
                break
    
    if json_file_path and os.path.exists(json_file_path):
        relationships = load_relationships_from_file(json_file_path)
        extracted_mappings = extract_enum_mappings(relationships)
        
        # 글로벌 TYPE_MAPPINGS 업데이트
        global TYPE_MAPPINGS
        TYPE_MAPPINGS.update(extracted_mappings)
        
        logger.info(f"enum 매핑 업데이트 완료: {len(extracted_mappings)} 타입")
    else:
        logger.warning("table_relationships.json 파일을 찾을 수 없습니다.")

def add_custom_enum_value(enum_type, value, table, column, description):
    """
    커스텀 enum 값을 추가합니다.
    
    Args:
        enum_type (str): enum 타입명
        value (str|int): enum 값
        table (str): 테이블명
        column (str): 컬럼명
        description (str): 설명
    """
    if enum_type not in CUSTOM_ENUM_VALUES:
        CUSTOM_ENUM_VALUES[enum_type] = {}
    
    CUSTOM_ENUM_VALUES[enum_type][str(value)] = {
        "table": table,
        "column": column,
        "description": description
    }
    
    logger.info(f"커스텀 enum 값 추가: {enum_type}.{value} -> {table}.{column}")

def get_custom_enum_value(enum_type, value):
    """
    커스텀 enum 값의 매핑 정보를 반환합니다.
    
    Args:
        enum_type (str): enum 타입명
        value (str|int): enum 값
        
    Returns:
        dict: 매핑 정보 또는 None
    """
    value_str = str(value)
    if enum_type in CUSTOM_ENUM_VALUES and value_str in CUSTOM_ENUM_VALUES[enum_type]:
        return CUSTOM_ENUM_VALUES[enum_type][value_str]
    return None

# 모듈 로드 시 자동으로 JSON에서 매핑 업데이트 시도
try:
    update_enum_mappings_from_json()
except Exception as e:
    logger.debug(f"자동 매핑 업데이트 실패 (정상): {e}")

# 사용 예시 함수들
def get_reward_type_table(reward_type_value):
    """RewardType 값에 해당하는 테이블 정보 반환"""
    return get_enum_value_mapping("RewardType", reward_type_value)

def get_cost_type_table(cost_type_value):
    """CostType 값에 해당하는 테이블 정보 반환"""
    return get_enum_value_mapping("CostType", cost_type_value)

def get_condition_type_table(condition_type_value):
    """ConditionType 값에 해당하는 테이블 정보 반환"""
    return get_enum_value_mapping("ConditionType", condition_type_value)