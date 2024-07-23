import os
import jieba
import re
import numpy as np
import pandas as pd
import win32com.client as winApp
import time

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

    def cal_score(self,forecast_result,custom_dict):
        search_result={}
        search_result_detail={}
        for test_ID,test_info in forecast_result.items():
            search_result[test_ID]={}
            search_result_detail[test_ID]={}
            for test_row,test_row_info in test_info.items():
                custom_contained=[]
                for custom_ID,custom_keys in custom_dict.items():
                    key_all=len(custom_keys)
                    key_contained=[]
                    custom_score=0.0
                    for key in custom_keys:
                        if test_row_info[0].find(key)>-1:
                            key_contained.append(key)
                    if len(key_contained)>0:
                        custom_contained.append(custom_ID)
                    custom_score=len(key_contained)/key_all
                    search_result_detail[test_ID]["Row{}-{}".format(test_row,custom_ID)]=[key_contained,custom_score]
                raw_sentence = test_row_info[0]
                search_result[test_ID][test_row] = [raw_sentence, custom_contained]
        return search_result,search_result_detail

    def cal_comb(self,custom_dict,timestamp,th=0.6):
        forecast_result_P,forecast_result_N,forecast_result_detail=self.cal_similarity(th)
        search_result, search_result_detail=self.cal_score(forecast_result_N,custom_dict)
        Title=["股票代码","原始句序","原始句文",r"与标准\个性化准则相似程度",r"符合标准\个性化准则"]
        final_table = []
        for testID in forecast_result_detail.keys():
            forecast_result_P_table=forecast_result_P[testID]
            search_result_table=search_result[testID]
            for row_id,row_info in forecast_result_P_table.items():
                final_table.append([testID,row_id,row_info[0],row_info[1],row_info[2]])
            for row_id,row_info in search_result_table.items():
                final_table.append([testID,row_id,row_info[0],None,row_info[1]])
        df=pd.DataFrame(final_table,columns=Title)
        df=df.sort_values(by=['股票代码','原始句序'],ascending=True)
        df.to_csv(r"Output\{}\PartII_收入政策相似度_{}.txt".format(timestamp,timestamp),sep="|",encoding="utf-8",index=False)
        self.__transform(df,timestamp)

    def __groupDF(self,df,key_col,sort_cols):
        df_Group={}
        all_keys=df[key_col].unique()
        for key in all_keys:
            key_details=df[df[key_col]==key].copy()
            key_details=key_details.sort_values(by=sort_cols,ascending=True)
            df_Group[key]=key_details
        return df_Group

    def __transform(self,df,timestamp):
        df_Group=self.__groupDF(df,"股票代码",["原始句序"])
        xlApp = winApp.Dispatch('Excel.Application')
        xlApp.Visible = False
        xlApp.ScreenUpdating = False
        xlApp.DisplayAlerts = False
        result_wb=xlApp.Workbooks.Open(os.path.join(os.path.dirname(__file__),r"Config\收入政策分类结果_Template.xlsx"), UpdateLinks=False)
        result_sht=result_wb.Worksheets['Sheet1']
        current_row=result_sht.Range("A1048576").End(-4162).Row+1
        for code,details in df_Group.items():
            t1=time.time()
            result_sht.Cells(current_row, 1).Value = code
            current_loc = 1
            current_loc_dict = {}
            for row_id in range(details.shape[0]):
                row_index = int(details.iloc[row_id, 1])
                row_txt = str(details.iloc[row_id, 2])
                row_txt = re.sub('\n', "", row_txt)
                row_txt_length = len(row_txt)
                row_sim = details.iloc[row_id, 4]
                if isinstance(row_sim, list):
                    temp = [0, 0]
                    raw_txt_all = result_sht.Cells(current_row, 3).Value
                    if len(row_sim) == 0:
                        typ = "N"
                    else:
                        typ = "C"
                    if raw_txt_all is None:
                        result_sht.Cells(current_row, 3).Value = row_txt + "。"
                        temp[0] = current_loc
                        current_loc = row_txt_length + 1
                        temp[1] = current_loc
                        current_loc_dict[row_index] = {"loc": temp, "type": typ}
                    else:
                        result_sht.Cells(current_row, 3).Value = raw_txt_all + row_txt + "。"
                        temp[0] = current_loc + 1
                        current_loc = current_loc + row_txt_length + 1
                        temp[1] = current_loc
                        current_loc_dict[row_index] = {"loc": temp, "type": typ}
                    for cus in row_sim:
                        if cus == "Custom1":
                            raw_txt_cus = result_sht.Cells(current_row, 13).Value
                            if raw_txt_cus is None:
                                result_sht.Cells(current_row, 13).Value = row_txt + "。"
                            else:
                                result_sht.Cells(current_row, 13).Value = raw_txt_cus + row_txt + "。"
                        elif cus == "Custom2":
                            raw_txt_cus = result_sht.Cells(current_row, 14).Value
                            if raw_txt_cus is None:
                                result_sht.Cells(current_row, 14).Value = row_txt + "。"
                            else:
                                result_sht.Cells(current_row, 14).Value = raw_txt_cus + row_txt + "。"
                        elif cus == "Custom3":
                            raw_txt_cus = result_sht.Cells(current_row, 15).Value
                            if raw_txt_cus is None:
                                result_sht.Cells(current_row, 15).Value = row_txt + "。"
                            else:
                                result_sht.Cells(current_row, 15).Value = raw_txt_cus + row_txt + "。"
                        elif cus == "Custom4":
                            raw_txt_cus = result_sht.Cells(current_row, 16).Value
                            if raw_txt_cus is None:
                                result_sht.Cells(current_row, 16).Value = row_txt + "。"
                            else:
                                result_sht.Cells(current_row, 16).Value = raw_txt_cus + row_txt + "。"
                        elif cus == "Custom5":
                            raw_txt_cus = result_sht.Cells(current_row, 17).Value
                            if raw_txt_cus is None:
                                result_sht.Cells(current_row, 17).Value = row_txt + "。"
                            else:
                                result_sht.Cells(current_row, 17).Value = raw_txt_cus + row_txt + "。"
                        elif cus == "Custom6":
                            raw_txt_cus = result_sht.Cells(current_row, 18).Value
                            if raw_txt_cus is None:
                                result_sht.Cells(current_row, 18).Value = row_txt + "。"
                            else:
                                result_sht.Cells(current_row, 18).Value = raw_txt_cus + row_txt + "。"
                        else:
                            pass
                else:
                    temp = [0, 0]
                    raw_txt_all = result_sht.Cells(current_row, 3).Value
                    if raw_txt_all is None:
                        result_sht.Cells(current_row, 3).Value = row_txt + "。"
                        temp[0] = current_loc
                        current_loc = row_txt_length + 1
                        temp[1] = current_loc
                        current_loc_dict[row_index] = {"loc": temp, "type": "S"}
                    else:
                        result_sht.Cells(current_row, 3).Value = raw_txt_all + row_txt + "。"
                        temp[0] = current_loc + 1
                        current_loc = current_loc + row_txt_length + 1
                        temp[1] = current_loc
                        current_loc_dict[row_index] = {"loc": temp, "type": "S"}
                    if row_sim == "Standard1":
                        raw_txt_std = result_sht.Cells(current_row, 4).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 4).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 4).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard2":
                        raw_txt_std = result_sht.Cells(current_row, 5).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 5).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 5).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard3":
                        raw_txt_std = result_sht.Cells(current_row, 6).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 6).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 6).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard4":
                        raw_txt_std = result_sht.Cells(current_row, 7).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 7).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 7).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard5":
                        raw_txt_std = result_sht.Cells(current_row, 8).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 8).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 8).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard6":
                        raw_txt_std = result_sht.Cells(current_row, 9).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 9).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 9).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard7":
                        raw_txt_std = result_sht.Cells(current_row, 10).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 10).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 10).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard8":
                        raw_txt_std = result_sht.Cells(current_row, 11).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 11).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 11).Value = raw_txt_std + row_txt + "。"
                    elif row_sim == "Standard9":
                        raw_txt_std = result_sht.Cells(current_row, 12).Value
                        if raw_txt_std is None:
                            result_sht.Cells(current_row, 12).Value = row_txt + "。"
                        else:
                            result_sht.Cells(current_row, 12).Value = raw_txt_std + row_txt + "。"
                    else:
                        pass
            for loc_id, loc_info in current_loc_dict.items():
                s, e = loc_info['loc']
                t = loc_info["type"]
                if t == "C":
                    result_sht.Cells(current_row, 3).GetCharacters(int(s), int(e)).Font.color = -1003520
                elif t == "S":
                    result_sht.Cells(current_row, 3).GetCharacters(int(s), int(e)).Font.color = -16776961
                else:
                    pass
            pd.set_option('display.max_columns',None)
            current_row += 1
            t2=time.time()
        result_wb.SaveAs(os.path.join(os.path.dirname(__file__), r"Output\{}\收入政策分类_最终底稿_{}.xlsx".format(timestamp,timestamp)))
        result_wb.Close(False)
        for wb in xlApp.Workbooks:
            wb.Close(False)
        xlApp.Quit()
