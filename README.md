# hwp_to_csv
아래 한글 문서로 작성된 신청서의 내용을 취합하여 dataframe을 생성한 뒤, csv로 저장하는 프로세스

## packages
사용한 패키지 중 olefile 로 새롭게 설치가 필요하다.
```python
pip install olefile
```
각 패키지의 역할은 다음과 같다.
- olefile은 아래 한글 파일을 읽는 역할
- os는 폴더 안에 있는 파일들의 리스트를 읽는 역할
- pandas는 뽑아낸 데이터를 데이터 프레임에 담아 csv 파일로 저장하는 역할
```python
import olefile
import os
import pandas as pd
```
## 과정
1. 폴더 내에 있는 모든 신청서 파일들을 읽어와 데이터에 담기 위해서 olefile로 읽은 신청서의 데이터 구조를 파악했다.
<img src="./sample_data/sample.jpg" width="450" height="300"><img src="https://user-images.githubusercontent.com/70126055/236431098-1c0590d6-9707-4bd1-9567-019e76cad0f0.png" width="450" height="250">

2. 파악한 구조를 바탕으로 뽑아내고 싶은 데이터를 리스트에 담아 내보내는 Application 클래스를 생성했다. (일부 생략)
```python
class Application():
  def __init__(self,path):
    self.text = self.get_text(path)
    self.datas = self.get_datas()
    
  # 아래 한글 문서를 text로 받기 
  def get_text(self,path):
    # 아래 한글 파일 불러오기기
    f = olefile.OleFileIO(path)
    encoded_txt = f.openstream("PrvText").read() # 1페이지 이내 분량이라 미리보기 뷰로도 충분히 내용을 가져올 수 있음
    text = encoded_txt.decode("utf-16",errors="ignore") # 디코딩

    # text 정제
    # 표의 cell은 <>로 감싸여 있음, 표의 내용만 가져오기 위해 다음과 같이 정제함
    text = text.replace(" ","").replace("><","\t").replace(">","") 
    text = text.split("\r\n")
    text = [t.replace("<","").split("\t") for t in text if len(t) > 0 and t[0] == "<"] 
    return text

  # 문서에서 데이터를 뽑아내기
  def get_datas(self):
    # '성명', '소속', '휴대폰' 등의 데이터 위치를 파악해 해당 내용만 뽑아 저장
    datas = []
    datas.append(self.text[0][1]) # 홍길동
    datas.append(self.text[1][1]) # ㅇㅇ대학교
    datas.append(self.text[2][2]) # 010-0000-0000	
    
    # 같은 행에 포함된 내용들을 한꺼번에 확인
    datas += [""]+[False]*3
    for param in self.text[3]:
      # ■ 혹은 ▣ 등으로 바뀌었다면 체크한 것으로 해당 문자로 입력함
      if param[0] != "□":
      
        # 참가형태는 3가지 중 하나만 선택 가능하기 때문에 우선 빈칸으로 저장하고 
        # 동일한 문자가 왔을 때 입력해주는 방식을 채택함
        if "학술대회+Workshop" in param:
          datas[3]="학술대회+Workshop"
        elif "학술대회" in param:
          datas[3]="학술대회"
        elif "Workshop" in param:
          datas[3]="Workshop"
        
        # 식사의 경우는 각각 만찬, 조식, 중식을 체크하거나 안할 수 있음으로 3가지로 나누어 TF로 저장함
        elif "만찬" in param: # 동일 문자가 포함된 경우 각 문자가 의미하는 위치에 True로 저장
          datas[4]=True
        elif "조식" in param:
          datas[5]=True
        elif "중식" in param:
          datas[6]=True    
    return datas
```
3. os.listdir 함수를 통해 폴더 내 신청서.hwp 문서 이름들을 받아 dataframe에 데이터를 저장해 csv로 저장했다.
```python
# 파일명 리스트
file_list = [name for name in os.listdir('.') if name[-3:]=="hwp"]

# 문서 내 데이터
datas = []
for path in file_list:
  data = Application(path)
  datas.append(data.datas)

# dataframe 에 담아 csv 혹은 엑셀로 저장
cols = ['성명', '소속', '휴대폰', '참가형태', '만찬', '조식', '중식']
df = pd.DataFrame(datas, columns=cols)
df.to_csv("sample.csv")
```

## 느낀점/배운점
- 데이터 파악을 잘한다면 문서 자동화도 쉽게 만들 수 있다!
- hwp 문서의 내용이 길 경우에는 BodyText/Section 읽어 파싱해야한다. [(아래한글 문서 읽기)](https://code-angie.tistory.com/49)
- 나중에 메일로 들어온 신청서를 자동으로 분류해서 내용을 뽑아낼 수 있도록 작업해보고 싶다.
