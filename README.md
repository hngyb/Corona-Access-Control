# QR코드 기반의 전자출입명부 자동화 시스템 개발
코로나19의 확산에 따라 도입된 QR코드 기반 전자출입명부를 자동화하는 시스템입니다. Python의 OpenCv와 pyzbar 라이브러리를 활용하여 QR코드를 스캔하고, Pandas 라이브러리를 활용하여 스캔된 정보와 입력한 정보를 엑셀 파일(xlsx)에 자동으로 저장합니다.
## 프로젝트 기간
2020.07.30 - 2020.07.31
## 개발 배경
- 기존 사용하던 바코드 스캐너(SYMcode, NEW S950 PLUS)의 한글 인코딩 문제로 스캔 정보를 자동으로 입력하고 관리하는 것의 어려움 발생.
- Python의 OpenCV와 pyzbar 라이브러리를 활용하면 스마트폰 카메라로 바코드 스캐너를 대체할 수 있음을 파악.
- Python의 Pandas 라이브러리를 활용하면 데이터 관리가 용이함을 파악.
- Python을 통해 QR코드를 스캔하고 전자출입명부 데이터를 관리하는 프로그램을 개발하기로 결정.
## 개발 과정
1. 스마트폰 카메라와 PC를 연결할 애플리케이션 탐색: iVCam
2. OpenCV와 pyzbar를 활용하여 영상 속 QR코드를 감지하고 디코딩
3. QR코드 정보를 utf-8로 디코딩
4. datetime을 통해 현재 시간을 방문 시간으로 자동 설정
5. Pandas의 Dataframe을 활용하여 QR코드 정보와 입력 정보를 관리 및 xlsx파일에 저장
6. Pyinstaller를 사용하여 .py 파일을 .exe 파일로 컨버팅
    - RecursionError: maximum recursion depth exceeded: spec파일에 "import sys\n sys.setrecursionlimit(5000)" 추가
    - Failed to load dynlib/dll libiconv.dll: libzbar-64.dll, libiconv.dll 파일을 실행파일과 같은 폴더에 넣기
## 사용 전 주의사항
1. 프로그램 실행 전, PC와 카메라를 우선 연결해주세요.
2. 스마트폰 카메라를 사용 시, iVCam 애플리케이션을 사용하여 연결해주세요.
3. "Corona-Access_Control.exe" 실행파일과 libzbar-64.dll, libiconv.dll 파일을 같은 폴더에 넣고 실행해주세요.
## 사용 방법
1. 이 프로그램은 현재 컴퓨터의 날짜와 시간으로 파일 이름과 방문 시간을 설정합니다.
2. 저장되는 파일 이름의 형식은 "YYYYMMDD_PC_PC번호.xlsx" 입니다.
3. 성도 한 명의 정보가 모두 입력될 때마다 저장이 이루어집니다.
4. __성도 정보 입력 중에는 저장된 파일을 열지 말아주세요. (오류가 발생합니다)__
5. __QR코드의 정해진 형식은 "교회/구역/이름/소속/연락처/" 입니다.__ QR코드 형식에 맞추어 코드를 수정해서 사용하세요.
6. 체온과 문진사항만 수동 입력하면, QR코드 정보와 방문 시간이 함께 엑셀에 자동 저장됩니다.
7. 문진사항 입력 시 enter를 입력하면 자동 'x' 처리됩니다.
8. 엑셀에 최종적으로 저장되는 형식은 __"교회/구역/이름/소속/연락처/체온/문진사항/방문시간"__ 입니다.
9. 문진사항 입력이 종료된 후, 'q'를 누르면 프로그램이 종료됩니다.
10. 문진사항 입력이 종료된 후, 'd'를 이전 기록이 삭제 됩니다.
11. 프로그램 시작 시에 이미 저장된 오늘 자의 파일이 있으면, 몇 번째 파일인지 묻는 메시지가 나옵니다.
12. 11번의 경우, 저장되는 파일 이름의 형식은 "YYYYMMDD_PC_PC번호_(num)" 입니다.
13. "QR Code Scanner"창에서 'q'를 입력하면 프로그램이 종료됩니다.
## 스크린샷
![20200731_134741_1111](https://user-images.githubusercontent.com/68726615/89001281-bb28e100-d334-11ea-9cc6-74eb1633f4b8.png)
   
![20200731_134741_2222](https://user-images.githubusercontent.com/68726615/89001282-bbc17780-d334-11ea-9de1-bb78aa3344cc.png)
   
![20200731_134741_3](https://user-images.githubusercontent.com/68726615/89001279-ba904a80-d334-11ea-8e0e-fed5e03969e4.png)
## exe 실행 파일 다운로드
[링크](https://drive.google.com/drive/folders/1YT46kycCBy-YfScHyM2EZOGodBU3_ewS?usp=sharing "")
