import os
import jieba
import re
import numpy as np
import pandas as pd
import win32com.client as winApp


def cleaning(s:str)->list:
    s = re.sub(r"\(.*?\)", "", s)
    s = re.sub(r"（.*?）", "", s)
    s = re.sub(r'[\W\s]', "", s)
    for i in "一二三四五六七八九十零":
        s = re.sub(i, "", s)
    for i in "0123456789":
        s = re.sub(i, "", s)
    s = re.sub('\n', "", s)
    ss=jieba.lcut(s)
    return ss

class WordLib(object):
    def __init__(self):
        self.raw = []
        self.lib=[]
        self.vecs=[]

    def add(self,ss:list):
        self.raw.append(ss)
        return None

    def generate_lib(self):
        for L in self.raw:
            for word in L:
                if not word in self.lib:
                    self.lib.append(word)
        return None

    def generate_vector(self,wordlist: list):
        dimension = len(self.lib)
        vector = [0] * dimension
        for word in wordlist:
            loc = self.lib.index(word)
            count = vector[loc]
            count += 1
            vector[loc] = count
        return vector

def cal_sim(vec_1,vec_2):
    sim = 0.0
    vec1 = np.array(vec_1)
    vec2 = np.array(vec_2)
    if not (np.linalg.norm(vec1) == 0) and not (np.linalg.norm(vec2) == 0):
        sim= vec1.dot(vec2) / (np.linalg.norm(vec1) * np.linalg.norm(vec2))
    return sim


if __name__=="__main__":
    #第二行 32列2023 33列2022 34列相似度
    pre_data = pd.read_excel(r"C:\Users\10174\Desktop\MyNote\Projects(Python)\About PDF\SimModel_R2\Config\Pre_金融工具政策.xlsx", sheet_name="Sheet1")
    Temp_data={}
    Mylib=WordLib()
    for row_id in range(pre_data.shape[0]):
        Code= pre_data.iloc[row_id, 0]
        Policy_2023=str(pre_data.iloc[row_id, 31])
        Policy_2022=str(pre_data.iloc[row_id, 32])
        Policy_2023_words=cleaning(Policy_2023)
        Policy_2022_words=cleaning(Policy_2022)
        Temp_data[Code] = {}
        Temp_data[Code]["2023P"] = Policy_2023
        Temp_data[Code]["2022P"] = Policy_2022
        Temp_data[Code]["2023P_w"] = Policy_2023_words
        Temp_data[Code]["2022P_w"] = Policy_2022_words
        Mylib.add(Policy_2023_words)
        Mylib.add(Policy_2022_words)
    del pre_data
    Mylib.generate_lib()
    for code,data in Temp_data.items():
        p_2023_w = Mylib.generate_vector(data["2023P_w"])
        p_2022_w = Mylib.generate_vector(data["2022P_w"])
        sim=cal_sim(p_2023_w,p_2022_w)
        Temp_data[code]["Sim"]=sim

    xlApp = winApp.Dispatch('Excel.Application')
    xlApp.Visible = False
    xlApp.ScreenUpdating = False
    xlApp.DisplayAlerts = False
    result_wb = xlApp.Workbooks.Open(os.path.join(os.path.dirname(__file__), r"Config\模板.xlsx"),
                                     UpdateLinks=False)
    result_sht = result_wb.Worksheets['Sheet1']
    current_row = result_sht.Range("A1048576").End(-4162).Row + 1

    for code, data in Temp_data.items():
        result_sht.Cells(current_row,1).Value=code
        result_sht.Cells(current_row, 32).Value = data["2023P"]
        result_sht.Cells(current_row, 33).Value = data["2022P"]
        result_sht.Cells(current_row, 34).Value = data["Sim"]

        current_row += 1
    result_wb.SaveAs(os.path.join(os.path.dirname(__file__), r"Output\最终底稿.xlsx"))
    result_wb.Close(False)
    for wb in xlApp.Workbooks:
        wb.Close(False)
    xlApp.Quit()