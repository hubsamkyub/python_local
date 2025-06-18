def apply_typecode_mappings(relationships, typecode_mappings):
    """
    RewardType 등 조건부 외래키 매핑 정보를 기존 관계 목록에 추가합니다.
    
    Args:
        relationships: 기존 관계 딕셔너리
        typecode_mappings: RewardType 등 조건 기반 매핑 정보 (list of dict)
    
    Returns:
        확장된 관계 딕셔너리
    """
    for item in typecode_mappings:
        src_tbl = item["source_table"]
        tgt_tbl = item["target_table"]
        tgt_col = item["target_column"]
        src_col = item["source_column"]
        rel = {
            "source_column": src_col,
            "target_column": tgt_col,
            "relation_type": "typecode_foreign_key",
            "filter_column": item.get("filter_column"),
            "filter_value": item.get("filter_value")
        }

        if src_tbl not in relationships:
            relationships[src_tbl] = {}
        if tgt_tbl not in relationships[src_tbl]:
            relationships[src_tbl][tgt_tbl] = []

        # 중복 방지
        if rel not in relationships[src_tbl][tgt_tbl]:
            relationships[src_tbl][tgt_tbl].append(rel)

    return relationships
