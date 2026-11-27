# safety_check_app/branches.py

HQ_BRANCH_MAP = {
    "test광주전남본부": [
        "영광지점",
        "나주지점",
        "광주지점",
        "목포지점",
        "순천지점",
        "여수지점",
    ],
    "test강원본부": [
        "춘천지점",
        "원주지점",
        "강릉지점",
    ],
    "test서울본부": [
        "강남지점",
        "강북지점",
    ],
}

def get_hq_list() -> list[str]:
    """전체 본부 목록을 반환합니다."""
    return list(HQ_BRANCH_MAP.keys())

def get_branches_by_hq(hq_name: str) -> list[str]:
    """선택된 본부에 소속된 지점 목록만 반환합니다."""
    return HQ_BRANCH_MAP.get(hq_name, [])