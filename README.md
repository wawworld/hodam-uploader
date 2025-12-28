# 🎓 호서대학교 Cando 시스템 자동 상담 입력 프로그램

호서대학교 Cando 시스템에 학생 상담 데이터를 자동으로 입력하는 프로그램입니다.

![Build Status](https://github.com/wawworld/hodam-uploader/workflows/Build%20Windows%20EXE/badge.svg)

## ✨ 주요 기능

- 📊 CSV 파일로 대량의 상담 데이터 일괄 처리
- 🤖 Playwright 기반 브라우저 자동화
- 🔄 시스템 Chrome 브라우저 자동 연결
- 📝 상세한 처리 결과 리포트 생성
- ⚡ 에러 발생 시 자동 복구 시도

## 📋 시스템 요구사항

### Windows
- Windows 10 이상
- Chrome 브라우저 설치 권장
- 인터넷 연결 필수

### 개발 환경 (소스코드 실행 시)
- Python 3.9 이상
- 필수 패키지: pandas, playwright, openpyxl

## 🚀 사용 방법

### 방법 1: EXE 파일 사용 (권장)

1. **다운로드**
   - [Releases 페이지](https://github.com/wawworld/hodam-uploader/releases)에서 최신 버전 다운로드
   - 또는 [Actions 탭](https://github.com/wawworld/hodam-uploader/actions)에서 최신 빌드 다운로드

2. **CSV 파일 준비**
   - 필수 컬럼: `학번`, `이름`, `상담일자`, `상담내용`
   - 예시:
```csv
   학번,이름,상담일자,제목,상담내용,상담분야,진로상태,취업상태
   20251001,홍길동,2025-01-15,진로상담,학생과 진로에 대해 상담함,진로,일반,관심
```

3. **실행**
```
   hodam_uploader.exe 파일을 더블클릭
   → CSV 파일 경로 입력
   → 브라우저에서 Cando 시스템 로그인
   → 자동 처리 시작
```

### 방법 2: Python 소스코드 실행

1. **저장소 클론**
```bash
   git clone https://github.com/wawworld/hodam-uploader.git
   cd hodam-uploader
```

2. **가상환경 생성 및 활성화**
```bash
   python -m venv venv
   
   # Windows
   venv\Scripts\activate
   
   # Mac/Linux
   source venv/bin/activate
```

3. **패키지 설치**
```bash
   pip install -r requirements.txt
   playwright install chromium
```

4. **실행**
```bash
   python hodam_uploader.py
```

## 📊 CSV 파일 형식

### 필수 컬럼
| 컬럼명 | 설명 | 예시 |
|--------|------|------|
| 학번 | 학생 학번 | 20251001 |
| 이름 | 학생 이름 | 홍길동 |
| 상담일자 | 상담 날짜 (YYYY-MM-DD) | 2025-01-15 |
| 상담내용 | 상담 내용 | 학생과 진로에 대해 상담함 |

### 선택 컬럼
| 컬럼명 | 설명 | 가능한 값 |
|--------|------|-----------|
| 제목 | 상담 제목 | 임의의 텍스트 |
| 상담분야 | 상담 분야 | 진로, 취업, 학습, 심리 등 |
| 상담구분 | 개인/집단 | 개인상담, 집단상담 |
| 상담시간_시 | 상담 시작 시간(시) | 0-23 |
| 상담시간_분 | 상담 시작 시간(분) | 0-59 |
| 진로상태 | 학생 진로 상태 | 일반, 관심, 중점 |
| 취업상태 | 학생 취업 상태 | 일반, 관심, 중점 |
| 학습상태 | 학생 학습 상태 | 일반, 관심, 중점 |
| 심리상태 | 학생 심리 상태 | 일반, 관심, 중점 |
| 전문상담의뢰 | 전문상담 의뢰 여부 | Y, N |
| 비공개설정 | 비공개 설정 여부 | Y, N |

### CSV 파일 예시
```csv
학번,이름,상담일자,제목,상담내용,상담분야,상담구분,진로상태,취업상태
20251001,홍길동,2025-01-15,진로상담,학생과 진로에 대해 상담을 진행함,진로,개인상담,일반,관심
20251002,김철수,2025-01-16,취업상담,취업 준비 과정에 대해 논의함,취업,개인상담,관심,관심
```

## 📁 출력 파일

프로그램 실행 후 다음 파일들이 생성됩니다:

- `cando_auto_log.txt`: 실행 로그
- `cando_auto_report_YYYYMMDD_HHMMSS.csv`: 처리 결과 리포트

## 🔧 문제 해결

### 브라우저를 찾을 수 없다는 오류
**증상**: `Executable doesn't exist` 오류

**해결방법**:
1. Chrome 브라우저를 설치하세요
2. 또는 명령 프롬프트에서 다음 실행:
```bash
   playwright install chromium
```

### 학생을 찾을 수 없다는 오류
**증상**: 특정 학생 검색 실패

**해결방법**:
1. 학번이 정확한지 확인
2. Cando 시스템에 해당 학생이 등록되어 있는지 확인
3. 지도학생 목록에 포함되어 있는지 확인

### CSV 파일 인코딩 오류
**증상**: 한글이 깨져서 보임

**해결방법**:
- CSV 파일을 UTF-8 인코딩으로 저장
- Excel에서: `다른 이름으로 저장` → `CSV UTF-8 (쉼표로 분리)`


## ⚠️ 중요: 예시 파일 사용 시 주의사항

제공된 예시 파일(`example.csv`, `example_simple.csv`)은 **테스트용 가상 데이터**입니다.

- 학번: `00000001`, `00000002` 등 명백히 가짜 번호
- 이름: 흔한 가명 사용
- 내용: 예시 목적의 임의 내용

**실제 사용 시**:
1. ⚠️ 예시 파일을 **절대 그대로 사용하지 마세요**
2. ✅ 실제 학생 데이터로 교체 필요
3. ✅ 개인정보 보호에 유의

**개인정보 보호**:
- CSV 파일에 민감한 개인정보가 포함될 수 있습니다
- 파일 관리에 주의하세요
- 사용 후 안전하게 삭제하세요


## 🛠️ 개발

### 프로젝트 구조
```
hodam-uploader/
├── hodam_uploader.py      # 메인 프로그램
├── requirements.txt        # Python 패키지 의존성
├── .github/
│   └── workflows/
│       └── build.yml       # GitHub Actions 빌드 설정
├── .gitignore
└── README.md
```

### 빌드 방법

GitHub Actions를 통해 자동으로 Windows EXE 파일이 빌드됩니다.

수동 빌드:
```bash
pip install pyinstaller
pyinstaller --onefile --name hodam_uploader hodam_uploader.py
```

## 📝 라이선스

이 프로젝트는 호서대학교 내부 사용을 목적으로 개발되었습니다.


## 📞 문의

문제가 발생하거나 개선 사항이 있으면 [Issues](https://github.com/wawworld/hodam-uploader/issues)에 등록해주세요.

## 📜 변경 이력

### v1.0.0 (2025-12-28)
- 초기 버전 출시
- CSV 파일 기반 자동 상담 입력
- 시스템 Chrome 브라우저 지원
- 처리 결과 리포트 생성
