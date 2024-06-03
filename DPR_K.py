# termainalで以下入力しておく。
#pip install streamlit
#pip install plotly-express
#pip install streamlit

import openpyxl
import pandas as pd
import plotly.express as px
import streamlit as st

st.set_page_config(page_title="Dashboard",
                   page_icon=":bar_chart:",
                  layout="wide"
)

#Strleamlitをフレッシュするたびにエクセルを毎回読み込むのを防ぐ。新規に読み込む場合は、@st,def,retun,"df = get_data_from_excel()"の前に＃つける。
#@st.cache_data  
#def get_data_from_excel():
df = pd.read_excel(
    io='output_K.xlsx',
    engine='openpyxl',
    sheet_name='Sheet1',
    skiprows=0,
    usecols='A:F',
    nrows=1000,
)
    #return df
#df = get_data_from_excel()

#st.dataframe(df)


#------Mainpage--------------
st.title(":bar_chart: Sapura/K2 Progress(2024yr) Dashboard")

# ----- sidebar----------
st.sidebar.header("Please filter here:")

location = st.sidebar.multiselect(
    "Select the Location:",
    options=df["Location"].unique(),
    default=df["Location"].unique().tolist()
)

category1 = st.sidebar.multiselect(
    "Select the Category1:",
    options=df["Category1"].unique(),
    default=df["Category1"].unique().tolist()
)

category2 = st.sidebar.multiselect(
    "Select the Category2:",
    options=df["Category2"].unique(),
    default=df["Category2"].unique().tolist()
)


date = st.sidebar.multiselect(
    "Select the Date:",
    options=df["Date"].unique(),
    default=df["Date"].unique().tolist()
)

df_selection = df.query(
    #"Date == @Date & Location == @Location & Category1 == @Category1 & Category2 == @Category2"
    "Date == @date & Location == @location & Category1 == @category1 & Category2 == @category2"
)

st.dataframe(df_selection)


# -24:00はDatetime型として存在しないので、23:59に変換する。(変換済み)
#df_selection['To'] = df_selection['To'].replace('24:00', '23:59:00')

# "-"が含まれる行が残っている場合、削除。時間計算時にエラーがでるため。(削除済み)
#df_selection = df_selection[(df_selection['From'] != '-') & (df_selection['To'] != '-') & (df_selection['Location'] != '-')]

# "From"と"To"の型をDatetime型に変換
df_selection['from'] = pd.to_datetime(df_selection['From'], format='%H:%M:%S')
df_selection['to'] = pd.to_datetime(df_selection['To'], format='%H:%M:%S')

# 時間差を計算して新しい列(hour)を追加
df_selection['hour'] = (df_selection['to'] - df_selection['from']).dt.total_seconds() / 3600

# 各Activityに対して要した時間をバーチャート（水平方向）にする。
df_location = df_selection[['hour', 'Location']]
df_location['Row'] = df_location.index
total_hours = df_location['hour'].sum()
fig_df_location = px.bar(
    df_location,
    x="hour",
    y="Location",
    orientation="h",
    title="<b>Hours for Activitis</b> <br>※平行で作業している時間も合計していることに注意。<br>サイドバーのフィルタリング機能を用い、各作業にかかった時間を分析するために使う。",
    color_discrete_sequence=["#0083B8"] * len(df_location),
    template="plotly_white",
    hover_data={"Row": True, "Location": True, "hour": True}
)
fig_df_location.update_traces(hovertemplate="Row: %{customdata}<br>Location: %{y}<br>Hours: %{x}<extra></extra>", customdata=df_location['Row'])

# 合計時間を表示する
for i, loc in enumerate(df_location["Location"]):
    fig_df_location.add_annotation(
        x=df_location["hour"].max() + 0.5,
        y=loc,
        #text=f"Total Hours: {df_location.groupby('Location')['hour'].sum()[loc]:.2f}",
        text=f"Total Hours: {df_location.groupby('Location')['hour'].sum()[loc]:.2f} ({(df_location.groupby('Location')['hour'].sum()[loc]/24):.2f} days)",
        showarrow=False,
        font=dict(size=20),
        align="left"
    )
    
fig_df_location.update_layout(
    width=1200,  # グラフの幅
    height=650  # グラフの高さ
)

st.plotly_chart(fig_df_location)

# バーチャート（水平方向）の作成終わり
# バーチャート（鉛直方向）の作成

# "date"と時刻を組み合わせる。
df_selection['from'] = pd.to_datetime(df_selection['Date'].astype(str) + ' ' + df_selection['From'].astype(str))
df_selection['to'] = pd.to_datetime(df_selection['Date'].astype(str) + ' ' + df_selection['To'].astype(str))

# ---------同じPFが連続し、一度しかDPRに出てこない場合はこれでよいが、再度PFが出てくると過剰に時間を計算してしまうため。このコードはVOID------
# "Location"が一致する行の最初と最後の行を抽出
#matched_rows = df_selection[df_selection.duplicated(subset=["Location"], keep=False)]  
#first_row = matched_rows.groupby(["Location"]).first().reset_index()  
#last_row = matched_rows.groupby(["Location"]).last().reset_index()  


# 新しいデータフレームを作成  最初と最後の行の情報を追加  
#df_PF = pd.DataFrame(columns=["from", "to", "Location"])  
#df_PF["from"] = first_row["from"]  
#df_PF["to"] = last_row["to"]  
#df_PF[["Location"]] = first_row[["Location"]]  
# ------------ここまで--------------------


# Locationの情報が連続する部分を一つのデータ群として処理
df_PF = pd.DataFrame(columns=["from", "to", "Location"])  
group_start = 0
current_loc = df_selection.at[0, 'Location']
for i, row in df_selection.iterrows():
    if row['Location'] != current_loc:
        # 新しいデータ群の最初の行から最後の行までの'from'列の最小値,新しいデータ群の最初の行から最後の行までの'to'列の最大値
        new_row = {'from': df_selection.loc[group_start:i-1, 'from'].min(),
                   'to': df_selection.loc[group_start:i-1, 'to'].max(),
                   'Location': df_selection.at[group_start, 'Location']}
        df_PF = pd.concat([df_PF, pd.DataFrame(new_row, index=[0])], ignore_index=True)
        
        # 次のデータ群の開始行を更新
        group_start = i
        current_loc = row['Location']

# 最後のデータ群の情報を抽出
last_row = {'from': df_selection.loc[group_start:len(df)-1, 'from'].min(),
            'to': df_selection.loc[group_start:len(df)-1, 'to'].max(),
            'Location': df_selection.at[len(df)-1, 'Location']}
df_PF = pd.concat([df_PF, pd.DataFrame(last_row, index=[0])], ignore_index=True)


# 時間差を計算して新しい列(Critical_time)を追加
df_PF['Critical_day'] = (df_PF['to'] - df_PF['from']).dt.total_seconds() / 3600 /24

# Locationごとに、クリティカルとなった時間の合計値をバーチャートにする。

fig_df_PF = px.bar(
    df_PF,
    x="Location",
    y="Critical_day",
    title="<b>Critical day for PFs</b><br>※PF名が連続する場合、そのデータ群で最も早い時間と最も遅い時間を抽出して差を計算する。<br>サイドバーのフィルタリング機能によりデータを削っても、データ群の最早最遅時間が不変ならグラフの結果は変わらない。",
    color_discrete_sequence=["#0083B8"] * len(df_PF),
    template="plotly_white",
)

fig_df_PF.update_layout(
    xaxis=dict(tickmode="linear"),
    plot_bgcolor="rgba(0,0,0,0)",
    yaxis=(dict(showgrid=True, dtick=1)),
)

text=df_PF.groupby('Location')['Critical_day'].sum().apply(lambda x: f"Total Hours: {x:.2f}").tolist()

fig_df_PF.update_traces(text=text, textposition='none', textfont=dict(size=20))

fig_df_PF.update_layout(
    width=1000,  # グラフの幅
    height=650  # グラフの高さ
)
st.plotly_chart(fig_df_PF)
