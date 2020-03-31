import re
import pandas as pd
import math
import xlsxwriter


coll = pd.read_excel('text.xlsx',index_col=0)

article = 90507 #總文章處理數量

#分別為三個裝df、tf、以及tf_idf的dictionary
tf = {}
tf['銀行'] = {}
tf['信用卡'] = {}
tf['匯率'] = {}
tf['台積電'] = {}
tf['台灣'] = {}
tf['日本'] = {}

df = {}
df['銀行'] = {}
df['信用卡'] = {}
df['匯率'] = {}
df['台積電'] = {}
df['台灣'] = {}
df['日本'] = {}

tf_idf = {}
tf_idf['銀行'] = {}
tf_idf['信用卡'] = {}
tf_idf['匯率'] = {}
tf_idf['台積電'] = {}
tf_idf['台灣'] = {}
tf_idf['日本'] = {}

#移除英文數字及特殊符號
def pre_process(text):
    text = re.sub(r'[^\w]','',text)
    text = re.sub(r'[A-za-z0-9]','',text)
    return text

#切詞
def to_ngram(dict_text,ngram, data):
    # tmp_gram = {}
    # print(dict_text)
    for i in range(1,len(dict_text)+1):
        # ret_list = []
        for j in range(len(dict_text[i])-ngram):
            w = dict_text[i][j:j+ngram]
            if w in tf[data]:
                tf[data][w] += 1
            else:
                tf[data][w] = 1
    # return tmp_gram

#分類文章
def to_topic(all_text,*topics):
    tmp_text = {}
    count = 1
    for i in range(1,len(all_text)+1):
        for topic in topics:
            if (topic in all_text[i]):
                tmp_text[count] = all_text[i]
                count += 1
                break
    return tmp_text

#透過字詞是否有在文章出現，判定df
def df_count(dict_text, data):
    for w in tf[data]:
        for i in range(1,len(dict_text)+1):
            if dict_text[i].find(w)!= -1 :
                if w in df[data]:
                    df[data][w] += 1

                else:
                    df[data][w] = 1 
    for w in list(df[data]): 
        if df[data][w] < 15:
            del df[data][w]
            del tf[data][w]

#將低於tf一定數量之字詞刪除，節省記憶體
def delete_tf(data):
	for w in list(tf[data]):
		if tf[data][w] < 50:
			del tf[data][w]

#進行合併字詞處裡，合併原理為df誤差1%以內
def merge_df(data):
	for w in df[data]:
		for x in df[data]:
			# print(df[data][w])
			# print(df[data][x])
			if w == x:
				pass
			elif x.find(w) != -1:
				if (df[data][w]*0.99 <= df[data][x]) and (df[data][w]*1.01 >= df[data][x]):
					# print("a")
					df[data][w] = 0
					# del tf[data][w]
					# del df[data][w]
					# tf[data].pop(w)
			else:
				pass
	for w in list(df[data]):
		if df[data][w] == 0:
			del df[data][w]
			del tf[data][w]


#算出tf_idf
def tf_idf_(data):
	for w in list(tf[data]):
		x = tf[data].get(w)
		y = df[data].get(w)
		tf_idf[data][w] = (1 + (math.log(x,10)))*(math.log(article/y, 10))
		if tf_idf[data][w] < 9.8:
			del tf_idf[data][w]

#算出tf_idf後，將重質性高且數據相近的字詞刪除
def merge_tf_idf(data):
	for w in list(tf_idf[data]):
		for x in list(tf_idf[data]):
			if w == x:
				pass
			elif x.find(w) != -1:
				if (tf_idf[data][w]*0.99 <= tf_idf[data][x]) and (tf_idf[data][w]*1.01 >= tf_idf[data][x]):
					tf_idf[data][w] = 0

	for w in list(tf_idf[data]):
		if tf_idf[data][w] == 0:
			tf_idf[data].pop(w)




#讀取數據，將不同條件擷取的數據放置不同index中
#擷取條件為單一字詞，即主題字詞本身
all_text = {}
for index,value in coll.iloc[:].iterrows():
    all_text[index] = pre_process(value['標題'] + value['內容'])
key = ['銀行','信用卡', '匯率', '台積電', '台灣', '日本' ]
Data = {}
Data['銀行'] = {}
Data['信用卡'] = {}
Data['匯率'] = {}
Data['台積電'] = {}
Data['台灣'] = {}
Data['日本'] = {}
text = []

for i in range(len(key)):
    dic = to_topic(all_text, key[i]) #將有關鍵字(例:銀行)的文章選出來，成為一個主題的文章集
    text.append(dic)#一個主題的文章集
    # print(len(text[i]))
del all_text

num = 0
for data in Data: #6個主題
	for n in range(2,7):
		to_ngram(text[num], n, data) #將某主題的文章集以n gram切詞
	
	#切詞後，進行df計算、刪除多餘字詞等步驟
	delete_tf(data)
	df_count(text[num], data)
	merge_df(data)
	tf_idf_(data)
	merge_tf_idf(data)
	
	num += 1



#因先前皆用dictionary分類，故在合併資料、輸出到excel時需要轉成list append後才能將不同數據都貼上excel
excel_out = {}
for i in range(len(key)):
	excel_out[key[i]] =sorted(tf_idf[key[i]].items(), key=lambda d: d[1])
	excel_out[key[i]].reverse()
	# print(type(excel_out[key[i]]))
	# print(type(excel_out[key[i]][0]))
	for j in range(100):
		excel_out[key[i]][j] = list(excel_out[key[i]][j])
		for w in tf[key[i]]:
			if excel_out[key[i]][j][0] ==  w:
				excel_out[key[i]][j].append(tf[key[i]][w])
				excel_out[key[i]][j].append(df[key[i]][w])
				excel_out[key[i]][j].append(j+1)

writer = pd.ExcelWriter('hw1_result.xlsx', engine='xlsxwriter')

for i in range(len(key)):
	df[i] = pd.DataFrame(excel_out[key[i]])
	df[i].columns = ['keyword', 'tf_idf', 'tf', 'df', '排名']
	df[i].to_excel(writer, sheet_name=key[i])

writer.save()