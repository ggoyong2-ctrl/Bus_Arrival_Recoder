◎ 파일설명

1. 실행파일

      -> 윈도우 : Bus_Arrival_Recoder(v1.x.xx).exe

      -> macOS : Bus_Arrival_Recoder(v1.x.xx).app

2. 인증키 저장위치

      -> key.cfg (AES-128암호화)

3. 공동배차 확인용 노선정보 엑셀파일

      -> 서울시버스노선기본정보(yyyymmdd).xlsx (자동 다운로드 및 저장됨, 크롬브라우저 설치 필요)
   
      -> 오픈API의 데이터에 공동배차 참여 회사의 전체가 나오지 않기에 필요한 파일입니다. 즉, 서울시 API 관리 소홀이에요.

4. GitHub에 파이썬 코드를 오픈소스로 올려두었으니 각자의 필요에 맞게 코드를 자유롭게 수정하여 쓰셔도 됩니다.
 
      -> 깃허브주소 : https://github.com/ggoyong2-ctrl/Bus_Arrival_Recoder 접속 후 Python Code 폴더 진입
   
      -> 파이썬코드 : Bus_Arrival_Recoder(v1.x.xx).py

      -> 응용프로그램 아이콘 : icon.ico

      -> macOS용 아이콘 : icon.icns

      -> 프로그램창 좌상단, 작업표시줄 아이콘 : icon.png

5. 참고자료 : API매뉴얼 Word파일 4건, 버스정보목록 Excel파일 4건



◎ 인증키 발급 및 Open API 사용신청 방법

1. https://www.data.go.kr 에서 회원가입을 진행해주세요.

2. 아래 페이지에 각각 접속하여 API활용신청하면, 인증키가 발급됩니다. 
(인증키 확인방법 : 로그인 -> 마이페이지 -> 개인 API인증키)
발급된 인증키는 보통 매주 월요일에 사용승인됩니다.
(사용승인되기 전에는 발급된 인증키를 입력하더라도 사용 불가함.)

      -> 서울특별시_버스도착정보조회 서비스 https://www.data.go.kr/data/15000314/openapi.do

      -> 서울특별시_정류소정보조회 서비스 https://www.data.go.kr/data/15000303/openapi.do

      -> 서울특별시_버스위치정보조회 서비스 https://www.data.go.kr/data/15000332/openapi.do

      -> 서울특별시_노선정보조회 서비스 https://www.data.go.kr/data/15000193/openapi.do

3. 각각의 API호출명 당, 각각 일일 1000트래픽이 부여되며, 호출명 당 1000건의 제한은 개인이 사용하기에는 다소 부족한 트래픽입니다.
다만, 1000건의 호출 이내로 60초 갱신기준 도착기록 시, 1개 노선 규모 정류장 2개의 24시간 데이터 수집은 충분히 가능합니다.
동봉해드린 파이썬 코드를 참고해서 본인만의 프로그램을 신규 작성하신분의 경우, 본인의 프로그램을 활용사례로 신청하시면 각각의 API호출명 당 일일 트래픽 10000건으로 상향됩니다. (10000건이면 개인 수준에서 필요한 정보를 수집하기에 적당한 수준임)

4. API호출 트래픽이 부족할 때, API트래픽 상향을 위한 활용사례 신청 방법

      -> https://www.data.go.kr 로그인  -> 마이페이지 -> 데이터 활용 -> Open API -> 활용신청 현황 -> 각각 활용신청 했던 서비스를 클릭 -> 활용신청 버튼 -> 본인의 활용 사례(앱 혹은 프로그램 스크린샷, 사이트주소 등 입력)



◎ 윈도우에서 파이썬 코드 컴파일하는 법 (선택사항, 코드를 수정해서 쓰실 분)

1. 동봉된 파이썬코드파일과 아이콘파일들이 위치하는 경로를 확인합니다. (예: D:\Bus_Arrival_Recoder\PythonCode)

2. 파이썬 설치 https://www.python.org/downloads/

3. 검색 -> cmd -> 프롬프트 창에 아래와 같이 입력합니다.

      -> pip install requests pandas openpyxl cryptography pyinstaller selenium webdriver-manager

4. 설치가 마무리되면, cmd 프롬프트 창에 아래와 같이 입력합니다. (예: D:\Bus_Arrival_Recoder\PythonCode)

      -> cd D:\Bus_Arrival_Recoder\PythonCode

      -> pyinstaller --icon="Icon.ico" --add-data "Icon.ico;." --add-data "Icon.png;." --windowed --onefile --name "Bus_Arrival_Recoder(v1.x.xx)" Bus_Arrival_Recoder(v1.x.xx).py

5. 파이썬 코드파일 위치의 dist폴더에 실행파일이 생성됩니다.



◎ macOS 에서 파이썬 코드 컴파일하는 법 (선택사항, 코드를 수정해서 쓰실 분)

1. 파이썬(Python) 설치하기

      -> 웹사이트 접속 : 파이썬 공식 다운로드 페이지(https://www.python.org/downloads/macos/에) 접속합니다.

      -> 파일 다운로드 : 화면에 보이는 노란색 버튼 Download Python 3.xx.x를 클릭합니다.

      -> 설치 진행 : 다운로드된 .pkg 파일을 더블 클릭하여 실행합니다. '계속'과 '동의'를 눌러 설치를 완료합니다.

      -> 터미널 창 실행 : 맥 화면 우측 상단의 돋보기(Spotlight)를 누르고 "터미널" 혹은 "Terminal"을 검색해 실행합니다. (흰색이나 검은색 글자 창이 뜹니다.)

      -> 설치 확인 : 터미널 창에 아래 명령어를 입력하고 엔터를 칩니다. 버전 숫자가 나오면 성공입니다.

      -> 터미널 창 입력문구 : python3 --version

2. 필수 부품(라이브러리) 설치하기

      -> 프로그램이 작동하는 데 필요한 부속들을 설치합니다. 터미널 창에 아래 명령어들을 한 줄씩 복사해서 붙여넣고 엔터를 누르세요.

      -> 패키지 관리 도구 업데이트 : python3 -m pip install --upgrade pip

      -> 필수 라이브러리 일괄 설치 : pip3 install requests pandas openpyxl cryptography selenium webdriver-manager Pillow pyinstaller

      -> 필수 라이브러리 설치 : xcode-select --install
 
3. 빌드용 폴더 준비하기

      -> 컴파일을 시작하기 전에 필요한 파일들을 한곳에 깔끔하게 모읍니다.

      -> 바탕화면에 새 폴더를 만들고 이름을 BusApp(예)으로 짓습니다.

      -> 이 폴더 안에 아래 3가지 파일을 복사해서 넣습니다.

         ●	Bus_Arrival_Recoder(v1.x.xx).py (파이썬 소스 코드)

         ●	icon.icns (맥용 시스템 아이콘 파일)

         ●	icon.png (프로그램 내부 화면 표시용 아이콘)

4. 터미널 창에서 폴더로 이동하기

      -> 터미널 창에 cd 를 입력합니다. (cd 다음에 한 칸 띄우기 필수!)

      -> 바탕화면에 만든 BusApp(예) 폴더를 마우스로 잡아서 터미널 창 안으로 드래그하여 떨어뜨립니다.

      -> 경로가 자동으로 입력되면 엔터를 누릅니다. (cd /Users/사용자명/Desktop/BusApp)

5. 앱 빌드 명령어 입력하기 (최종 단계)

      -> 이제 진짜 맥용 앱을 만드는 명령어입니다. 아래 내용을 통째로 복사해서 터미널 창에 붙여넣고 엔터를 누르세요. 

      -> 터미널 창 입력문구 : pyinstaller --icon="Icon.icns" --add-data "Icon.png:." --add-data "Icon.icns:." --windowed --onefile --name "Bus_Arrival_Recoder(v1.x.xx)" Bus_Arrival_Recoder(v1.x.xx).py

      -> 기다림 : 터미널 창에 영어 글자들이 빠르게 올라갑니다. 약 1~2분 정도 기다리세요.

      -> 완료 : 마지막 줄에 successfully라는 단어가 보이면 완성된 것입니다.

6. 결과물 확인 및 실행 (중요)

      -> BusApp(예) 폴더 안에 dist라는 폴더가 생겼습니다. 그 안으로 들어갑니다.

      -> Bus_Arrival_Recoder라는 이름의 앱 파일이 보입니다.

      -> 구글 크롬이 설치(설치파일 googlechrome.dmg)되어야 공동배차 확인용 노선정보 엑셀파일을 다운로드 할 수 있습니다. 

      -> 구글 크롬 설치 오류 시, 터미털 창 입력문구 : xattr -cr /Applications/Google\ Chrome.app

7. 맥에서 처음 실행할 때 주의사항 (보안 정책)

      -> 애플 공식 인증 앱이 아니므로 처음엔 실행이 안 될 수 있습니다. 이때는 다음과 같이 하세요.

      -> 앱을 마우스 오른쪽 버튼(또는 두 손가락 클릭)으로 누르고 [열기]를 선택합니다.

      -> 확인 창이 뜨면 다시 한번 [열기]를 누릅니다.

      -> 한 번만 이렇게 실행해 주면 그 뒤로는 그냥 더블 클릭만으로 잘 실행됩니다.

      -> 도움이 되셨길 바랍니다! 혹시 가이드 내용 중 추가로 궁금한 점이 있으시면 말씀해 주세요.
