## 프로젝트 소개

 ESC/POS 프린트
 Windows 환경에서 사용가능
 
## 기술 스택

|Category| - |
| --- | --- |
|Language|Python v3.11.11|
|Framework|Flask v3.1.0|



### 사용법

Execute the following lines to properly clone and run the project.  

COMMAND
- $ conda create -n ["env"] python=3.11.11
- $ conda activate ["env"]
- $ pip install -r requirements.txt
- $ python esc_pos.py

RELEASE 
- `.exe` 파일을 **릴리즈** 목록에서 다운로드 하여 실행 하 실 수 있습니다.


실행파일 만들기
- $ pyinstaller --noconfirm --onefile --windowed --icon=printer.ico --add-data "templates;templates" --add-data "printer.ico;." --hidden-import=plyer.platforms.win.notification --hidden-import=backports --hidden-import=backports.tarfile --name "Coleslaw Printer" esc_pos.py
