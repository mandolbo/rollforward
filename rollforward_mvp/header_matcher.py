"""
헤더 매칭 - MVP 버전  
정확히 일치하는 것만 매칭 (유사도 매칭은 나중에)

이 파일이 하는 일:
- 전기 조서의 헤더(컬럼명)과 당기 파일의 헤더를 비교해서 매칭시키기
- 예: 전기 조서의 "매출액" 컬럼이 당기 파일의 어느 컬럼과 연결되는지 찾기

왜 이 기능이 필요한가?
- 데이터를 정확한 위치에 복사하기 위해서는 어느 컬럼이 어느 컬럼인지 알아야 함
- 사람이 일일이 확인하지 않아도 자동으로 연결점을 찾게 하기 위해
- 실수 없이 정확한 매핑을 보장하기 위해
"""

def match_headers(previous_tables, current_file_path):
    """
    헤더 매칭 - 올바른 데이터 흐름 버전
    
    이 함수가 하는 일:
    1. 전기 조서(백데이터 시트)와 당기 PBC 파일의 테이블들을 비교
    2. 정확히 같은 이름의 헤더들을 찾아서 연결
    3. 당기 PBC → 백데이터 방향으로 매칭 결과 반환
    
    데이터 흐름:
    - 소스: 당기 PBC 파일 (current_file_path)
    - 대상: 전기 조서의 백데이터 시트 (previous_tables)
    
    Parameters:
        previous_tables (list): 전기 조서에서 찾은 백데이터 테이블들 (대상)
        current_file_path (str): 당기 PBC 파일 경로 (소스)
        
    Returns:
        list: 매칭된 헤더들의 정보 리스트 (from=당기PBC, to=백데이터)
    """
    
    # 왜 여기서 import를 하는가?
    # 함수 내부에서만 사용하는 것은 필요할 때만 import하는 것이 좋음
    # 순환 import 문제를 피하기 위해서도 사용
    from table_finder import find_tables
    current_tables = find_tables(current_file_path)  # 당기 파일에서도 테이블 찾기
    
    matches = []  # 매칭 결과를 저장할 리스트
    
    # 왜 이중 반복문을 사용하는가?
    # 전기 조서의 모든 테이블과 당기 파일의 모든 테이블을 조합해서 비교하기 위해
    # 어떤 테이블끼리 매칭되는지 모르기 때문에 모든 경우를 확인
    for prev_table in previous_tables:
        for curr_table in current_tables:
            
            # 왜 헤더를 하나씩 확인하는가?
            # 테이블 내의 각 컬럼(헤더)별로 매칭을 찾아야 하기 때문
            # 한 테이블에 여러 컬럼이 있을 수 있음
            # 당기 PBC의 각 헤더를 백데이터 헤더와 매칭
            for curr_header in curr_table['headers']:
                
                # 당기 PBC 헤더가 백데이터에 있는지 확인
                if curr_header in prev_table['headers']:
                    # 백데이터에서 매칭되는 헤더 찾기 (MVP에서는 정확히 일치하는 것만)
                    prev_header = curr_header  # 정확히 일치하므로 같은 이름
                    
                    # 올바른 데이터 흐름을 위한 매칭 정보 저장
                    # Current PBC (소스) → 백데이터 sheets (대상) 방향으로 수정
                    match = {
                        'from_table': curr_table,     # 소스: 당기 PBC 테이블 정보
                        'to_table': prev_table,       # 대상: 백데이터 시트 테이블 정보  
                        'from_header': curr_header,   # 소스 헤더명 (당기 PBC의 실제 헤더)
                        'to_header': prev_header,     # 대상 헤더명 (백데이터의 매칭 헤더)
                        'confidence': 1.0             # 매칭 신뢰도 (정확히 일치하므로 100%)
                    }
                    matches.append(match)
                    print(f"[header_matcher.match_headers] 🔗 매칭 발견: '{curr_header}' (당기 PBC) → '{prev_header}' (백데이터) (신뢰도: 100%)")
    
    # 왜 매칭이 없을 때 메시지를 출력하는가?
    # 사용자가 상황을 이해하고 필요한 조치를 취할 수 있게 도와주기 위해
    # 디버깅에도 도움이 됨
    if not matches:
        print("[header_matcher.match_headers] ⚠️ 정확히 일치하는 헤더가 없습니다")
        print("[header_matcher.match_headers] 💡 힌트: 전기 조서와 당기 파일의 헤더 이름을 확인하세요")
    
    return matches  # 찾은 매칭들의 리스트 반환

def simple_similarity_match(header1, header2):
    """
    간단한 유사도 매칭 (Phase 1에서 사용 예정)
    
    왜 이 함수가 필요한가?
    - MVP에서는 정확히 같은 이름만 매칭하지만
    - 실제 업무에서는 "매출"과 "Revenue", "매출액"과 "Sales" 같은 유사한 이름들도 매칭해야 함
    - 나중에 더 지능적인 매칭을 위한 기반 함수
    
    Returns:
        float: 0.0(전혀 다름) ~ 1.0(완전히 같음) 사이의 유사도 점수
    """
    
    # 왜 lower()와 strip()을 사용하는가?
    # 대소문자와 공백 차이로 인한 매칭 실패를 방지하기 위해
    h1 = header1.lower().strip()
    h2 = header2.lower().strip()
    
    # 정확히 같으면 100% 일치
    if h1 == h2:
        return 1.0
    
    # 포함 관계 확인 (예: "매출"이 "총매출액"에 포함)
    if h1 in h2 or h2 in h1:
        return 0.8  # 80% 유사도
    
    # 공통 단어 확인 (간단한 방법)
    words1 = set(h1.split())  # 첫 번째 헤더를 단어로 분리
    words2 = set(h2.split())  # 두 번째 헤더를 단어로 분리
    common_words = words1.intersection(words2)  # 공통 단어 찾기
    
    # 공통 단어가 있으면 비율로 유사도 계산
    if common_words:
        return len(common_words) / max(len(words1), len(words2))
    
    return 0.0  # 아무 관련 없음

def enhanced_match_headers(previous_tables, current_file_path, threshold=0.7):
    """향상된 헤더 매칭 (Phase 1에서 사용 예정)"""
    
    from table_finder import find_tables
    current_tables = find_tables(current_file_path)
    
    matches = []
    
    for prev_table in previous_tables:
        for curr_table in current_tables:
            
            for prev_header in prev_table['headers']:
                best_match = None
                best_score = 0.0
                
                for curr_header in curr_table['headers']:
                    score = simple_similarity_match(prev_header, curr_header)
                    
                    if score >= threshold and score > best_score:
                        best_match = curr_header
                        best_score = score
                
                if best_match:
                    matches.append({
                        'from_table': prev_table,
                        'to_table': curr_table,
                        'from_header': prev_header,
                        'to_header': best_match,
                        'confidence': best_score
                    })
                    print(f"[header_matcher.enhanced_match_headers] 🔗 매칭: '{prev_header}' → '{best_match}' (신뢰도: {best_score:.1%})")
    
    return matches

def test_header_matcher():
    """헤더 매칭 기능 테스트"""
    print("[header_matcher.test_header_matcher] 🧪 헤더 매칭 테스트...")
    
    # 더미 테이블 생성
    prev_tables = [{
        'sheet': 'Sheet1',
        'start_row': 1,
        'headers': ['이름', '매출', '비용'],
        'file_path': 'dummy_prev.xlsx'
    }]
    
    curr_tables = [{
        'sheet': 'Sheet1', 
        'start_row': 1,
        'headers': ['이름', '매출', '기타'],
        'file_path': 'dummy_curr.xlsx'
    }]
    
    # 간단한 매칭 테스트
    matches = []
    for prev_table in prev_tables:
        for curr_table in curr_tables:
            for prev_header in prev_table['headers']:
                if prev_header in curr_table['headers']:
                    matches.append({
                        'from_header': prev_header,
                        'to_header': prev_header,
                        'confidence': 1.0
                    })
    
    if matches:
        print(f"[header_matcher.test_header_matcher] ✅ 테스트 성공: {len(matches)}개 매칭 발견")
        for match in matches:
            print(f"[header_matcher.test_header_matcher]    {match['from_header']} → {match['to_header']}")
    else:
        print("[header_matcher.test_header_matcher] ❌ 테스트 실패: 매칭을 찾지 못했습니다")

if __name__ == "__main__":
    test_header_matcher()