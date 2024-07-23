import os
import jieba
import re
import numpy as np
import pandas as pd
import win32com.client as winApp


class MyBag(object):
    def __init__(self):
        self.bag={}

    def __cleaning(self,s:str)->str:
        s = re.sub(r"\(.*?\)", "", s)
        s = re.sub(r"（.*?）", "", s)
        s = re.sub(r'[\W\s]', "", s)
        for i in "一二三四五六七八九十零":
            s=re.sub(i,"",s)
        for i in "0123456789":
            s=re.sub(i,"",s)
        s=re.sub('\n',"",s)
        return s

    def add(self, tag: str, doc: str):
        self.bag[tag] = {}
        sentences_raw = doc.split("。")
        sentences_raw.pop()
        sentences_cleansed = [self.__cleaning(x) for x in sentences_raw]
        sentences_cut = [jieba.lcut(x) for x in sentences_cleansed]
        self.bag[tag]["raw"] = sentences_raw
        self.bag[tag]["processed"] = sentences_cut
        return None

class Model_Similarity(object):
    def __init__(self,tag_bag,test_bag):
        self.Word_Library=[]
        self.__generate_lib(tag_bag)
        self.__generate_lib(test_bag)
        self.Tag_bag=self.__generate_vectors(tag_bag)
        self.Test_bag=self.__generate_vectors(test_bag)

    def __generate_lib(self,bag):
        for tag,info in bag.bag.items():
            cut_list = info['processed']
            for row in cut_list:
                for word in row:
                    if not word in self.Word_Library:
                        self.Word_Library.append(word)
        return None

    def __generate_vectors(self,bag):
        new_bag=bag
        for tag,info in new_bag.bag.items():
            cut_list = info['processed']
            new_bag.bag[tag]['vector']=[self.generate_vector(x) for x in cut_list]
        return new_bag

    def generate_vector(self,wordlist: list):
        dimension = len(self.Word_Library)
        vector = [0] * dimension
        for word in wordlist:
            loc = self.Word_Library.index(word)
            count = vector[loc]
            count += 1
            vector[loc] = count
        return vector

    def calculate_probability(self,vec_1: list, vec_2: list):
        prob = 0.0
        vec1 = np.array(vec_1)
        vec2 = np.array(vec_2)
        if not (np.linalg.norm(vec1) == 0) and not (np.linalg.norm(vec2) == 0):
            prob = vec1.dot(vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2))
        return prob

    def cal_similarity(self,th=0.6):
        forecast_result_positive={}
        forecast_result_negative={}
        forecast_result_detail={}
        for test_ID,test_info in self.Test_bag.bag.items():
            forecast_result_positive[test_ID]={}
            forecast_result_negative[test_ID]={}
            forecast_result_detail[test_ID]={}
            test_vectors=test_info['vector']
            test_row=1
            for test_vector in test_vectors:
                max_prob=0.0
                most_sim=None
                for standard_ID,standard_info in self.Tag_bag.bag.items():
                    standard_vectors=standard_info['vector']
                    standard_row = 1
                    for standard_vector in standard_vectors:
                        prob=self.calculate_probability(test_vector,standard_vector)
                        forecast_result_detail[test_ID]["{}-{}_{}-{}".format(test_ID, test_row,standard_ID,standard_row)] = prob
                        if prob>max_prob:
                            max_prob=prob
                            most_sim=standard_ID
                        standard_row+=1
                raw_sentence=self.Test_bag.bag[test_ID]['raw'][test_row-1]
                if max_prob>=0.6:
                    forecast_result_positive[test_ID][test_row]=[raw_sentence,max_prob,most_sim]
                else:
                    forecast_result_negative[test_ID][test_row]=[raw_sentence,max_prob,most_sim]
                test_row+=1
        return forecast_result_positive,forecast_result_negative,forecast_result_detail

    def __groupDF(self,df,key_col,sort_cols):
        df_Group={}
        all_keys=df[key_col].unique()
        for key in all_keys:
            key_details=df[df[key_col]==key].copy()
            key_details=key_details.sort_values(by=sort_cols,ascending=True)
            df_Group[key]=key_details
        return df_Group
    def cal_comb(self,th=0.6):
        forecast_result_P,forecast_result_N,forecast_result_detail=self.cal_similarity(th)
        Title=["股票代码","原始句序","原始句文",r"与标准相似程度",r"符合标准"]
        final_table = []
        for testID in forecast_result_detail.keys():
            forecast_result_P_table=forecast_result_P[testID]
            for row_id,row_info in forecast_result_P_table.items():
                final_table.append([testID,row_id,row_info[0].replace("\n",""),row_info[1],row_info[2]])
        df=pd.DataFrame(final_table,columns=Title)
        df=df.sort_values(by=['股票代码','原始句序'],ascending=True)
        df.to_csv(r"Output\相似度文件.txt",sep="|",encoding="utf-8",index=False)
        df_Group = self.__groupDF(df, "股票代码", ["原始句序"])
        return df_Group


if __name__=="__main__":
    df_standard = pd.read_csv(r"Config\标准政策.txt", sep="|", encoding="utf-8")
    bag_standard = MyBag()
    for row_id in range(df_standard.shape[0]):
        s_id = df_standard.iloc[row_id, 0]
        s_txt = df_standard.iloc[row_id, 1]
        bag_standard.add(s_id, s_txt)

    df_sample = pd.read_excel(r"C:\Users\MY598RL\Desktop\result_0514_全量政策\SZ.xlsx", sheet_name="Sheet1")
    bag_sample = MyBag()
    for row_id in range(df_sample.shape[0]):
        test_id = df_sample.iloc[row_id, 0]
        test_txt = df_sample.iloc[row_id, 1]
        bag_sample.add(str(test_id), str(test_txt))

    model = Model_Similarity(bag_standard, bag_sample)
    df_G=model.cal_comb(0.6)

    xlApp = winApp.Dispatch('Excel.Application')
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    result_wb = xlApp.Workbooks.Open(os.path.join(os.path.dirname(__file__), r"Config\标准底稿.xlsx"),
                                     UpdateLinks=False)
    result_sht = result_wb.Worksheets['Sheet1']
    current_row = result_sht.Range("A1048576").End(-4162).Row + 1

    for row_id in range(df_sample.shape[0]):
        code = df_sample.iloc[row_id, 0]
        C_txt = str(df_sample.iloc[row_id, 1]).replace("\n","")
        result_sht.Cells(current_row, 1).Value = code
        result_sht.Cells(current_row, 3).Value = C_txt
        C_txt_Ss=C_txt.split("。")
        s_txt=""
        for sentence in C_txt_Ss:
            if sentence.find("研发支出的归集范围")>-1:
               s_txt=s_txt+sentence+"。"
        result_sht.Cells(current_row, 5).Value = s_txt

        if code in df_G.keys():
            details=df_G[code]
            G_txt=""
            I_txt=""
            sim_policy= {"K":[],"L":[],"M":[],"N":[],"O":[]}
            for row_id in range(details.shape[0]):
                row_index = int(details.iloc[row_id, 1])
                row_txt = str(details.iloc[row_id, 2])
                row_sim_val = float(details.iloc[row_id, 3])
                row_sim_res = str(details.iloc[row_id, 4])
                sim_policy[row_sim_res].append(row_sim_val)
                if row_sim_res in ["K","L","M"]:
                    G_txt = G_txt + "。" + re.sub('\n', "", row_txt)
                elif row_sim_res in ["N","0"]:
                    I_txt = I_txt + "。" + re.sub('\n', "", row_txt)
                else:
                    pass
            if len(G_txt)>1:
                G_txt=G_txt[1:]
            if len(I_txt)>1:
                I_txt=I_txt[1:]
            sum1_range=[]
            sum2_range=[]
            for p,v in sim_policy.items():
                if len(v)>0:
                    val = np.average(v)
                    if p=="K" and len(v)>0:
                        result_sht.Cells(current_row, 11).Value = val
                        sum1_range.append(val)
                    elif p=="L" and len(v)>0:
                        result_sht.Cells(current_row, 12).Value = val
                        sum1_range.append(val)
                    elif p == "M" and len(v)>0:
                        result_sht.Cells(current_row, 13).Value = val
                        sum1_range.append(val)
                    elif p == "N" and len(v)>0:
                        result_sht.Cells(current_row, 14).Value = val
                        sum2_range.append(val)
                    elif p == "O" and len(v)>0:
                        result_sht.Cells(current_row, 15).Value = val
                        sum2_range.append(val)
                    else:
                        pass
            result_sht.Cells(current_row, 7).Value = G_txt
            result_sht.Cells(current_row, 8).Value =np.average(sum1_range)
            result_sht.Cells(current_row, 9).Value = I_txt
            result_sht.Cells(current_row, 10).Value =np.average(sum2_range)
        current_row += 1
    result_wb.SaveAs(os.path.join(os.path.dirname(__file__), r"Output\最终底稿.xlsx"))
    result_wb.Close(False)
    for wb in xlApp.Workbooks:
        wb.Close(False)
    xlApp.Quit()