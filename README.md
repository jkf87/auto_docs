# auto_docs
automatically generates a PDF file by putting data (CSV) into a Korean document (HWP).

출결 서류가 너무 많아서 HWP 자동화 코드를 생성했습니다.

환경설정이 귀찮으실듯 하여 exe파일도 함께 배포합니다
[파일 다운로드](https://m100000577-my.sharepoint.com/:u:/g/personal/fitz87_samyang_es_kr/EUFpZnvOe6lAodMAoWdmB2MBWapNNeOGXTDqEnHEzK2Vsw?e=xYDLZz)


## 업데이트
- 기존 파일명을 이름으로 지정하던 코드를 수정 -> 첫번째 필드로 이름을 지정
- (24.2.18) 이미지 파일을 사용할 수 있도록 수정, 폴더안에 번호 순으로 정렬되어 들어감, PDF, HWP 두가지 형태로 저장할 수 있도록 옵션 제공
- (24. 3.4) 데이터가 비어 있는 경우 모든 데이터가 첫번째 자리 소수점을 찍는 버그를 수정, xls, xlsx파일 지원, exe파일 버전 0.03v으로 업데이트
  
## 의존성(Dependencies)
pandas: Used for data manipulation and analysis. [Pandas Documentation](https://pypi.org/project/pandas/)

pywin32: Provides access to many of the Windows APIs from Python. [PyWin32 Documentation](https://pypi.org/project/pywin32/)

Pillow : [docs](https://pillow.readthedocs.io/en/stable/)

### 설치방법(Installation Steps)
이 프로젝트를 시작하려면 먼저 다음 명령을 사용하여 리포지토리를 복제합니다:

To get started with this project, first clone the repository using the following command:
```
git clone https://github.com/jkf87/auto_docs.git
```

다음 명령을 사용하여 필요한 패키지를 설치합니다

Install the required packages using the following command:

```
pip install -r requirements.txt
```
# 기본 사용법
CSV 파일 내에 있는 첫번째 행의 이름과 HWP 파일 내에 누름틀(Ctrl+K+E)의 이름과 갯수를 맞춰주세요.

예시)

첫행과 누름틀 이름을 매칭시켜주세요.

CSV 파일 내의 첫행 이름(데이터 라벨)

![image](https://github.com/jkf87/auto_docs/assets/28688071/4bc3ca04-1341-4311-8883-f6475a771eed)

HWP 파일 내의 누름틀 이름(빨간색으로 된 부분)

![image](https://github.com/jkf87/auto_docs/assets/28688071/a3f4fc08-6c54-42a4-9afd-31f28981b056)


이후 파이썬 파일(main.py)를 실행해주세요.

1.데이터 파일(CSV)을 첨부

![image](https://github.com/jkf87/auto_docs/assets/28688071/097decf9-7cfe-483d-a787-b4bab66e1f41)

2.문서 파일(HWP)를 첨부

![image](https://github.com/jkf87/auto_docs/assets/28688071/7dca175f-06c1-47e2-910c-55f842c24a72)

3.모두 허용 누르기

![image](https://github.com/jkf87/auto_docs/assets/28688071/cbd25279-a70f-49e8-90ec-a76a2343d49c)

4.문서 생성

![출결서류생성](https://github.com/jkf87/auto_docs/assets/28688071/a74f5f6b-6533-4f2e-aeda-cee32ba47abe)

5.결과물 확인하기

![image](https://github.com/jkf87/auto_docs/assets/28688071/cd53c287-2971-42a3-a1f6-26d2050a93a2)

#  최신 버전 사용법 ( 유튜브 )

[![image](https://i.ytimg.com/vi/GuEdVQKFFE8/hqdefault.jpg?sqp=-oaymwEjCNACELwBSFryq4qpAxUIARUAAAAAGAElAADIQj0AgKJDeAE=&rs=AOn4CLACIrsXE9a9DnvRkWCc0W0JpEZF6Q)](https://youtu.be/GuEdVQKFFE8?si=7VHo5drFDYHSpKik)







 

































