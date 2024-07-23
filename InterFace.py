import pandas as pd
import datetime
import time
import os
from PartI import My_Extractor
from PartII import MyBag,Model_Similarity


class Interface_Part1(object):
    def __init__(self,tolerate=50):
        self.t=tolerate

    def main(self):
        timestamp = "{}".format(datetime.datetime.now().strftime('%Y%m%d%H%M%S'))
        os.makedirs(r"Output\{}".format(timestamp))

        with open(r"PartI_Config.config", "r+", encoding="utf-8") as f1:
            content = f1.read()
            part1_config = content.split(sep="\n")
            part1_config.pop()
            part1_config = [x.split(sep="|") for x in part1_config]

        with open(r"Output\{}\PartI_收入政策摘取_{}.txt".format(timestamp, timestamp), "a+", encoding="utf-8") as f:
            f.write("Code|Policy|Time|Remark\n")
        for row_id in range(len(part1_config)):
            if row_id==0:
                continue
            t1 = time.time()
            BU0 = part1_config[row_id][0]
            p=part1_config[row_id][1]
            status=part1_config[row_id][2]
            if status=="Done":
                continue
            me = My_Extractor(p, 50)
            try:
                _, policy = me.main()
                Remark = "Success"
            except Exception as e:
                policy = ""
                Remark = "Failed=>{}".format(e)
            t2 = time.time()
            with open(r"Output\{}\PartI_收入政策摘取_{}.txt".format(timestamp, timestamp), "a+", encoding="utf-8") as f:
                f.write("{}|{}|{}s|{}\n".format(BU0, policy, t2 - t1, Remark))
            print("At {}>>>Spend {}s".format(BU0, t2 - t1))


class Interface_Part2(object):
    def __init__(self):
        self.paths={
                    "S":r"Config\标准收入准则.xlsx",
                    "C":r"Config\个性化收入政策.xlsx"
        }

    def main(self,timestamp,tolerate=0.6):
        df_standard = pd.read_excel(self.paths["S"], sheet_name="Sheet1")
        bag_standard = MyBag()
        for row_id in range(df_standard.shape[0]):
            s_id = df_standard.iloc[row_id, 0]
            s_txt = df_standard.iloc[row_id, 1]
            bag_standard.add(s_id, s_txt)

        df_custom = pd.read_excel(self.paths["C"], sheet_name="Sheet1")
        dict_custom = {}
        for row_id in range(df_custom.shape[0]):
            c_id = df_custom.iloc[row_id, 0]
            c_txt = df_custom.iloc[row_id, 1]
            dict_custom[c_id] = set(str(c_txt).split(';'))

        df_sample = pd.read_csv(r"Output\{}\PartI_收入政策摘取_{}.txt".format(timestamp,timestamp), sep="|",encoding="utf-8")
        bag_sample = MyBag()
        for row_id in range(df_sample.shape[0]):
            test_id = df_sample.iloc[row_id, 0]
            test_txt = df_sample.iloc[row_id, 1]
            if str(test_txt)[2] == "、":
                bag_sample.add(str(test_id), str(test_txt))

        model = Model_Similarity(bag_standard, bag_sample)
        model.cal_comb(dict_custom, timestamp, tolerate)

if __name__=="__main__":
    # i1=Interface_Part1(50)
    # i1.main()

    i2=Interface_Part2()
    i2.main("20240517000000",0.6)
