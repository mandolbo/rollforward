"""
ChatGPT 디버깅용 정보 수집기
개발 중 문제 발생 시 이 파일을 실행하고 결과를 ChatGPT에 업로드하세요!
"""

# 왜 이 라이브러리들을 import하는가?
import os         # 파일과 폴더 정보를 가져오기 위한 라이브러리
import sys        # Python 시스템 정보를 가져오기 위한 라이브러리 (버전, 경로 등)
import traceback  # 오류 발생 시 상세한 오류 정보를 가져오기 위한 라이브러리
from datetime import datetime  # 현재 날짜와 시간을 기록하기 위한 라이브러리
import platform   # 운영체제 정보를 가져오기 위한 라이브러리 (Windows, Mac, Linux)

def collect_all_info():
    """프로젝트 모든 정보를 ChatGPT 업로드용으로 수집"""
    
    print("🔍 디버깅 정보 수집 중...")
    
    info = []
    info.append("=" * 80)
    info.append("🐛 롤포워딩 프로젝트 디버깅 정보")
    info.append(f"📅 수집 시간: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    info.append("=" * 80)
    
    # 1. 시스템 정보
    info.append("\n📊 1. 시스템 정보")
    info.append("-" * 40)
    info.append(f"Python 버전: {sys.version}")
    info.append(f"운영체제: {platform.system()} {platform.release()}")
    info.append(f"프로세서: {platform.processor()}")
    info.append(f"현재 작업 디렉토리: {os.getcwd()}")
    
    # 2. 프로젝트 구조
    info.append("\n📁 2. 프로젝트 구조")
    info.append("-" * 40)
    for root, dirs, files in os.walk("."):
        level = root.replace(".", "").count(os.sep)
        indent = " " * 2 * level
        info.append(f"{indent}{os.path.basename(root)}/")
        subindent = " " * 2 * (level + 1)
        for file in files:
            if file.endswith(('.py', '.xlsx', '.txt', '.md')):
                try:
                    file_size = os.path.getsize(os.path.join(root, file))
                    info.append(f"{subindent}{file} ({file_size} bytes)")
                except:
                    info.append(f"{subindent}{file} (크기 불명)")
    
    # 3. Python 파일 내용
    info.append("\n📄 3. Python 파일 내용")
    info.append("-" * 40)
    
    python_files = []
    for root, dirs, files in os.walk("."):
        for file in files:
            if file.endswith('.py') and not file.startswith('__'):
                python_files.append(os.path.join(root, file))
    
    # 파일 우선순위 (중요한 것부터)
    priority_order = ['main.py', 'table_finder.py', 'header_matcher.py', 'file_updater.py']
    sorted_files = []
    
    # 우선순위 파일들 먼저
    for priority in priority_order:
        for file_path in python_files:
            if priority in file_path:
                sorted_files.append(file_path)
                break
    
    # 나머지 파일들
    for file_path in python_files:
        if file_path not in sorted_files:
            sorted_files.append(file_path)
    
    for file_path in sorted_files:
        try:
            info.append(f"\n📝 === {file_path} ===")
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
                lines = content.split('\n')
                
                # 토큰 제한 고려: 파일이 너무 크면 일부만
                if len(lines) > 100:
                    info.append("⚠️ 파일이 큼: 처음 50줄과 마지막 50줄만 표시")
                    info.extend([f"{i+1:3d}: {line}" for i, line in enumerate(lines[:50])])
                    info.append("... (중간 생략) ...")
                    info.extend([f"{i+len(lines)-50+1:3d}: {line}" for i, line in enumerate(lines[-50:])])
                else:
                    info.extend([f"{i+1:3d}: {line}" for i, line in enumerate(lines)])
                    
        except Exception as e:
            info.append(f"❌ 파일 읽기 실패: {e}")
    
    # 4. 최근 에러 로그 (있다면)
    info.append("\n🚨 4. 최근 에러 정보")
    info.append("-" * 40)
    
    try:
        # 가장 최근에 발생한 에러 정보 수집
        if hasattr(sys, 'last_traceback') and sys.last_traceback:
            info.append("마지막 에러 트레이스백:")
            info.extend(traceback.format_tb(sys.last_traceback))
        else:
            info.append("현재 활성 에러 없음")
    except:
        info.append("에러 정보 수집 실패")
    
    # 5. 설치된 패키지 (requirements.txt나 import 에러 체크용)
    info.append("\n📦 5. 필요 패키지 상태")  
    info.append("-" * 40)
    
    required_packages = ['openpyxl', 'pandas', 'tkinter']
    for package in required_packages:
        try:
            if package == 'tkinter':
                import tkinter
                info.append(f"✅ {package}: 설치됨")
            else:
                __import__(package)
                info.append(f"✅ {package}: 설치됨")
        except ImportError:
            if package == 'tkinter':
                info.append(f"❌ {package}: 설치 필요 (Python 재설치 시 tkinter 포함 선택)")
            else:
                info.append(f"❌ {package}: 설치 필요 (pip install {package})")
    
    # 6. 샘플 데이터 구조 (있다면)
    info.append("\n📊 6. 테스트 파일 정보")
    info.append("-" * 40)
    
    if os.path.exists("test_files"):
        for root, dirs, files in os.walk("test_files"):
            for file in files:
                if file.endswith('.xlsx'):
                    file_path = os.path.join(root, file)
                    try:
                        file_size = os.path.getsize(file_path)
                        info.append(f"📄 {file_path} ({file_size} bytes)")
                    except:
                        info.append(f"📄 {file_path} (크기 불명)")
    else:
        info.append("⚠️ test_files 폴더가 없습니다")
    
    # 7. 환경 변수 (일부만)
    info.append("\n🔧 7. 환경 정보")
    info.append("-" * 40)
    env_vars = ['PATH', 'PYTHONPATH', 'HOME', 'USER']
    for var in env_vars:
        value = os.environ.get(var, '없음')
        if len(value) > 100:
            value = value[:100] + "...(생략)"
        info.append(f"{var}: {value}")
    
    # 결과 저장
    output_file = f"debug_info_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
    
    try:
        with open(output_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(info))
        
        print(f"✅ 디버깅 정보 수집 완료!")
        print(f"📁 파일 생성: {output_file}")
        print(f"📊 총 {len(info)}줄의 정보가 수집되었습니다")
        
    except Exception as e:
        print(f"❌ 파일 저장 실패: {e}")
        print("📋 콘솔 출력으로 대신 표시:")
        print("\n".join(info[:50]))  # 처음 50줄만 표시
        return None
    
    print("\n🤖 ChatGPT 사용 방법:")
    print("1. 위 파일을 ChatGPT에 업로드")
    print("2. '이 롤포워딩 프로젝트의 문제를 분석해주세요' 라고 질문")
    print("3. 구체적인 에러나 문제 상황도 함께 설명")
    print("4. 예시: 'main.py를 실행했는데 ModuleNotFoundError가 발생해요'")
    
    return output_file

def quick_test():
    """빠른 테스트 - 각 모듈이 제대로 import되는지 확인"""
    print("🧪 빠른 테스트 시작...")
    
    tests = [
        ("openpyxl 패키지", lambda: __import__('openpyxl')),
        ("main.py 모듈", lambda: __import__('main')),
        ("table_finder.py", lambda: __import__('table_finder')),
        ("header_matcher.py", lambda: __import__('header_matcher')),
        ("file_updater.py", lambda: __import__('file_updater')),
    ]
    
    success_count = 0
    total_count = len(tests)
    
    for test_name, test_func in tests:
        try:
            test_func()
            print(f"✅ {test_name}: OK")
            success_count += 1
        except Exception as e:
            print(f"❌ {test_name}: {e}")
    
    print(f"🧪 빠른 테스트 완료: {success_count}/{total_count} 성공")
    
    if success_count == total_count:
        print("🎉 모든 테스트 통과! 기본 설정이 완료되었습니다.")
    else:
        print("⚠️ 일부 테스트 실패. debug_collector.py의 전체 정보 수집을 권장합니다.")

def check_sample_files():
    """샘플 파일 존재 여부 확인"""
    print("📂 샘플 파일 확인 중...")
    
    # 왜 특정 파일들을 확인하지 않는가?
    # 사용자가 직접 선택하는 방식으로 변경되었기 때문에
    # 더 이상 고정된 샘플 파일에 의존하지 않음
    print("📁 샘플 파일 확인을 건너뜁니다.")
    print("💡 파일 선택은 프로그램 실행 시 사용자가 직접 수행합니다.")
    return
    
    missing_files = []
    for file_path in required_files:
        if os.path.exists(file_path):
            print(f"✅ {file_path}: 존재")
        else:
            print(f"❌ {file_path}: 없음")
            missing_files.append(file_path)
    
    if missing_files:
        print(f"\n⚠️ {len(missing_files)}개 파일이 누락되었습니다:")
        for file in missing_files:
            print(f"   - {file}")
        print("\n💡 해결 방법:")
        print("   1. 샘플 Excel 파일들을 직접 생성하거나")
        print("   2. 기존 파일 경로를 main.py에서 수정하세요")
    else:
        print("🎉 모든 샘플 파일이 준비되었습니다!")

if __name__ == "__main__":
    print("🛠️ ChatGPT 디버깅 도구")
    print("1. 전체 정보 수집 (권장)")
    print("2. 빠른 테스트만")
    print("3. 샘플 파일 확인")
    print("4. 종료")
    
    while True:
        choice = input("\n선택 (1-4): ").strip()
        
        if choice == "1":
            collect_all_info()
            break
        elif choice == "2":
            quick_test()
            break
        elif choice == "3":
            check_sample_files()
            break
        elif choice == "4":
            print("👋 안녕히 가세요!")
            break
        else:
            print("❌ 잘못된 선택입니다. 1-4 중에서 선택하세요.")