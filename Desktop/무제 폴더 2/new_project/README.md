# 🔧 DSD Breaker 프로젝트

> Excel Add-in 기반 데이터 분석 및 자동화 도구 개발

## 📋 프로젝트 개요

DSD Breaker는 Excel Add-in(.xlam) 형태의 데이터 처리 도구로, 복잡한 데이터 구조를 분석하고 자동화하는 시스템입니다.

## 🎯 목표

- [ ] 기존 Excel Add-in 기능 분석
- [ ] Python 기반 개선 버전 개발
- [ ] 사용자 친화적 인터페이스 구축
- [ ] 자동화 기능 확장

## 📁 프로젝트 구조

```
dsd_breaker/
├── original/
│   └── DsDbreaker.xlam     # 원본 Excel Add-in
├── analysis/               # 기능 분석 결과
├── python_version/         # Python 구현 버전
└── documentation/          # 문서화
```

## 🚀 개발 진행 상황

### Phase 1: 분석 단계 ✅
- [x] 원본 파일 확보
- [x] 기능 분석 (XLAM 구조 분석 완료)
- [x] 요구사항 정의

### Phase 2: 개발 단계 ✅
- [x] Python 기반 재구현 (차트 생성 포함)
- [x] UI/UX 개선 (tkinter 기반 GUI)
- [x] 테스트 및 검증 (샘플 데이터 포함)

### Phase 3: 배포 단계 ✅
- [x] 문서화 완성 (README, 사용법)
- [x] 사용자 가이드 작성
- [x] 실행 스크립트 생성

## 🛠️ 기술 스택

- **분석**: Python, openpyxl, pandas
- **개발**: Python, tkinter/PyQt
- **테스트**: pytest, unittest
- **문서화**: Markdown, Sphinx

---

**시작일**: 2025년 6월 23일  
**완료일**: 2025년 6월 23일  
**상태**: ✅ 완성

## 🎯 주요 성과

### 📊 DSD Breaker Python Edition 완성 (2가지 버전)

#### 🔧 일반 데이터 분석 버전 (`dsd_breaker_concept.py`)
- **6가지 차트 타입** 지원 (막대/선/산점도/히스토그램/파이/박스)
- **한글 폰트 자동 설정** (크로스 플랫폼)
- **완전한 GUI 인터페이스** (tkinter 기반)
- **대용량 데이터 처리** (1000+ 행 최적화)
- **다양한 저장 형식** (Excel, CSV, PNG, PDF, SVG)

#### 🔍 DART 감사보고서 검증 버전 (`dsd_breaker_audit.py`)
- **HTML/Excel 파일 지원** (DART 감사보고서 특화)
- **재무제표 레벨 자동 감지** (들여쓰기 기반)
- **합계 검증 시스템** (보고값 vs 계산값)
- **교차 참조 확인** (테이블 간 일관성 검증)
- **오류 패턴 자동 탐지** (이상 데이터 감지)

### 🔍 원본 XLAM 분석 완료
- 121,950 bytes 파일 구조 분석
- VBA 프로젝트 (232,960 bytes) 확인
- 25개 내부 파일, 7개 이미지 리소스 분석
- Custom UI 및 관계 파일 분석

### 🚀 실행 환경
```bash
# 의존성 설치
pip install -r requirements.txt

# === 일반 데이터 분석 버전 ===
# 테스트 데이터 생성
python3 test_dsd_breaker.py
# 프로그램 실행
python3 launch_dsd_breaker.py

# === DART 감사보고서 검증 버전 ===
# 감사용 테스트 데이터 생성
python3 test_audit_features.py
# 감사 검증 도구 실행
python3 dsd_breaker_audit.py
```