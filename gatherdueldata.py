import streamlit as st
import openpyxl as ox
import pandas as pd


# データを読み込み、worksheet型と1次元dfのリストにする デッキリストだけは別でリストに

def datastopy(path):
    import datetime
    dueldatas_master = ox.load_workbook(path)
    dueldatas = dueldatas_master["シート1"]
    datas = []
    decks = []
    i=7
    while True:
        if dueldatas.cell(row=i,column=1).value is None:
            break
        datefront = dueldatas.cell(row=i,column=1).value
        decks.append(dueldatas.cell(row=i,column=2).value)
        duel = []
        for j in range(2,12):
            duel.append(dueldatas.cell(row=i,column=j).value)
        dfduel = pd.DataFrame({
            "デッキ":duel[0],
            "対戦数":duel[1],
            "先手":duel[2],
            "後手":duel[3],
            "先手勝ち":duel[4],
            "先手負け":duel[5],
            "後手勝ち":duel[6],
            "後手負け":duel[7],
            "先手勝率":duel[8],
            "後手勝率":duel[9]
        },index=[datefront])
        datas.append(dfduel)
        i += 1   
    return dueldatas_master,datas,decks

def pytodatas(dueldatas_master,deck,order,result):
    import datetime
    dueldatas = dueldatas_master["シート1"]
    datefront = datetime.datetime.now()
    datefront = "{}/{}/{}".format(datefront.year,datefront.month,datefront.day)
    i = 7
    cal = 1
    while True:
        dueldate = dueldatas.cell(row=i,column=1).value
        deckdataed = dueldatas.cell(row=i,column=2).value
        if dueldate is None:
            cal = 0
            break
        if dueldate == datefront and deckdataed  == deck:
            if order == "先手" and result == "勝ち":
                dueldatas.cell(row=i,column=1,value=datefront)
                changepoint = int(dueldatas.cell(row=i,column=3).value)
                dueldatas.cell(row=i,column=3,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=4).value)
                dueldatas.cell(row=i,column=4,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=6).value)
                dueldatas.cell(row=i,column=6,value=changepoint+1)
                break
            elif order == "後手" and result == "勝ち":
                dueldatas.cell(row=i,column=1,value=datefront)
                changepoint = int(dueldatas.cell(row=i,column=3).value)
                dueldatas.cell(row=i,column=3,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=5).value)
                dueldatas.cell(row=i,column=5,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=8).value)
                dueldatas.cell(row=i,column=8,value=changepoint+1)
                break
            elif order == "先手" and result == "負け":
                dueldatas.cell(row=i,column=1,value=datefront)
                changepoint = int(dueldatas.cell(row=i,column=3).value)
                dueldatas.cell(row=i,column=3,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=4).value)
                dueldatas.cell(row=i,column=4,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=7).value)
                dueldatas.cell(row=i,column=7,value=changepoint+1)
                break
            elif order == "後手" and result == "負け":
                dueldatas.cell(row=i,column=1,value=datefront)
                changepoint = int(dueldatas.cell(row=i,column=3).value)
                dueldatas.cell(row=i,column=3,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=5).value)
                dueldatas.cell(row=i,column=5,value=changepoint+1)
                changepoint = int(dueldatas.cell(row=i,column=9).value)
                dueldatas.cell(row=i,column=9,value=changepoint+1)
                break
        i += 1
    if cal == 0:
        if order == "先手" and result == "勝ち":
            dueldatas.cell(row=i,column=1,value=datefront)
            dueldatas.cell(row=i,column=2,value=deck)
            dueldatas.cell(row=i,column=3,value=1)
            dueldatas.cell(row=i,column=4,value=1)
            dueldatas.cell(row=i,column=5,value=0)
            dueldatas.cell(row=i,column=6,value=1)
            dueldatas.cell(row=i,column=7,value=0)
            dueldatas.cell(row=i,column=8,value=0)
            dueldatas.cell(row=i,column=9,value=0)
            
        elif order == "後手" and result == "勝ち":
            dueldatas.cell(row=i,column=1,value=datefront)
            dueldatas.cell(row=i,column=2,value=deck)
            dueldatas.cell(row=i,column=3,value=1)
            dueldatas.cell(row=i,column=4,value=0)
            dueldatas.cell(row=i,column=5,value=1)
            dueldatas.cell(row=i,column=6,value=0)
            dueldatas.cell(row=i,column=7,value=0)
            dueldatas.cell(row=i,column=8,value=1)
            dueldatas.cell(row=i,column=9,value=0)
            
            
            
        elif order == "先手" and result == "負け":
            dueldatas.cell(row=i,column=1,value=datefront)
            dueldatas.cell(row=i,column=2,value=deck)
            dueldatas.cell(row=i,column=3,value=1)
            dueldatas.cell(row=i,column=4,value=1)
            dueldatas.cell(row=i,column=5,value=0)
            dueldatas.cell(row=i,column=6,value=0)
            dueldatas.cell(row=i,column=7,value=1)
            dueldatas.cell(row=i,column=8,value=0)
            dueldatas.cell(row=i,column=9,value=0)
            
        elif order == "後手" and result == "負け":
            dueldatas.cell(row=i,column=1,value=datefront)
            dueldatas.cell(row=i,column=2,value=deck)
            dueldatas.cell(row=i,column=3,value=1)
            dueldatas.cell(row=i,column=4,value=0)
            dueldatas.cell(row=i,column=5,value=1)
            dueldatas.cell(row=i,column=6,value=0)
            dueldatas.cell(row=i,column=7,value=0)
            dueldatas.cell(row=i,column=8,value=0)
            dueldatas.cell(row=i,column=9,value=1)            
        
     

def pytoadditionaldata(df,dueldatas):
    for i in range(len(df)):
        first = df.iloc[i]["先手"]
        firstwin = df.iloc[i]["先手勝ち"]
        second = df.iloc[i]["後手"]
        secondwin = df.iloc[i]["後手勝ち"]
        if first != 0:
            rate_firstwin = round(firstwin/first,3)
            dueldatas.cell(row=i+7,column=10,value=rate_firstwin)
        if second != 0:    
            rate_secondwin = round(secondwin/second,3)
            dueldatas.cell(row=i+7,column=11,value=rate_secondwin)

def datas_init(dueldatas_master):
    dueldatas = dueldatas_master["シート1"]
    for row in dueldatas.iter_rows(min_row=7, min_col=1, max_row=600, max_col=11):
        for cell in row:
            cell.value = None

    



       
# ページレイアウト
st.set_page_config(
    page_title="DC 戦績記入,分析ツール",
    page_icon="🧊",
    layout="wide",
    initial_sidebar_state="collapsed"  ,
    menu_items={}
)

# 対戦データの読み込み、デッキ表示
dueldatas_master,datalist,deckdueled = datastopy("database_florting/dueldatas.xlsx")
dueldatas = dueldatas_master["シート1"]

 #　表示の調整　重複するデッキをはじいている  
deckdueled = set(deckdueled)
deckdueled = list(deckdueled)

if "decks" not in st.session_state:
    st.session_state.decks = deckdueled
    
deck_options = st.session_state.decks


# ここまでが読み込み　ここから動的な部分

st.title("DC 戦績記入,分析ツール")

# デッキ情報の取り出し　記入モジュール
newdeck = st.text_input("新しいデッキの追加")
if st.button("追加"):
    apd = 0
    for i in st.session_state.decks:
        if newdeck == i:
            st.write("そのデッキは追加されています")
            apd = 1
            break
    if apd == 0 and newdeck != "":
        deck_options.append(newdeck)

# 選択肢に実質的なプレイスホルダーを追加
reject = "ここに入力 元の文字は消さない"
if deck_options == []:
    deck_options.append(reject)
else:
    if deck_options[0] != reject:
        deck_options.insert(0,reject)
    
deck = st.selectbox("対戦したデッキを選んでください 直打ちで検索もできます",deck_options)
order = st.radio("先手後手を記入",("先手","後手"),horizontal=True)
result = st.radio("勝ち負けを記入",("勝ち","負け"),horizontal=True)

submit = st.button("結果を記入")
# submit により書き込み
if submit is True:
    if deck == reject:
        st.write("無効なデッキ名です")
    else:
        pytodatas(dueldatas_master,deck,order,result)
        dueldatas_master.save("database_florting/dueldatas.xlsx")
        st.button("データの同期") 
        
    
if datalist == []:
    st.write("データがありません。")
else:
    df = pd.concat(datalist)
    pytoadditionaldata(df,dueldatas)
    dueldatas_master.save("database_florting/dueldatas.xlsx")
    st.write(df)
    st.button("勝率を計算")

st.markdown("### 危険　全データの初期化")

check = st.checkbox("初期化しますか？")
check2 = st.checkbox("こうかいしませんね？")
if st.button("このボタンでデータが初期化") and check and check2:
    datas_init(dueldatas_master)
    dueldatas_master.save("database_florting/dueldatas.xlsx")
    st.write("データを初期化しました。　リロードすると反映されます")

    
   
    





