# Windows 11 실행파일(.exe) 빌드/배포

## 방법 1: GitHub Actions 자동 빌드 (권장)

1. GitHub 저장소의 `Actions` 탭으로 이동
2. `Build Windows EXE` 워크플로우 선택
3. `Run workflow` 클릭
4. 완료 후 실행 결과의 `Artifacts`에서 `SpendTagger-windows-exe` 다운로드
5. 내려받은 `SpendTagger.exe` 파일 하나만 친구에게 전달

## 방법 2: Windows PC에서 직접 빌드

사전 준비:
- Python 3.12 설치

명령:
```bat
build_windows_exe.bat
```

결과:
- `dist\SpendTagger.exe`

## 전달/실행 가이드

- 친구 PC에서 `SpendTagger.exe`를 더블클릭 실행
- SmartScreen 경고가 나오면 `추가 정보` -> `실행` 선택
- 결과 파일(엑셀/차트)은 실행한 작업 폴더에 생성됨
