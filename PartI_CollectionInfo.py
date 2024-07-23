import os
import re
import pdfplumber


class MyPdfBase(object):
    def __init__(self,p):
        self.File=pdfplumber.open(p)
        self.Bu=self.__get_name(p)

    def __get_name(self,p):
        name=os.path.basename(p)
        name=name.replace(".pdf","")
        name=name.replace(".PDF","")
        return name

    def _get_index(self,page,target:str):
        page_index=None
        page_msg=page.extract_text().replace(" ","")
        if_found=page_msg.find(target)
        if if_found>-1:
            page_index = page.page_number
        return page_index

    def _get_location(self,page,target:str):
        page_list = page.extract_words()
        marking_location = [None,None,None,None]
        for item in page_list:
            item_text = item["text"]
            item_text = item_text.replace(" ", "")
            if item_text == target:
                marking_location[0] = item["x0"]
                marking_location[1] = item["top"]
                marking_location[2] = item["x1"]
                marking_location[3] = item["bottom"]
        marking_location =tuple(marking_location)
        return marking_location

class My_Extractor(MyPdfBase):
    def __init__(self,p,tolerate=50):
        super().__init__(p)
        self.markings={
                    "layer1_marking":r"第十节[\s]*财务报告",
                    "layer2_marking":"重要会计政策及会计估计",
                    "begin_marking":r"[0-9][0-9]、[\s]*收入",
                    "end_marking":r"{}、.*"
                        }
        self.tolerate=tolerate

    def main(self):
        m_l1_found = None
        m_l2_found = None
        begin_found ={"Page":None,"Details":[]}
        index = None
        end_found = {"Page":None,"Details":[]}
        current_text=[]
        final_result=""
        for page in self.File.pages:
            if page.page_number <= self.tolerate:
                continue
            if not m_l1_found:
                if len(page.search(self.markings["layer1_marking"])) > 0:
                    m_l1_found = page.page_number
            if m_l1_found and not m_l2_found:
                if len(page.search(self.markings["layer2_marking"])) > 0:
                    m_l2_found = page.page_number
            if m_l1_found and m_l2_found and begin_found["Page"] is None:
                begin_found["Details"] = page.search(pattern=self.markings["begin_marking"])
                if len(begin_found["Details"])==1:
                    begin_found["Page"]=page.page_number
                    index=re.search(r"[0-9][0-9]",begin_found["Details"][0]["text"]).group()
            if m_l1_found and m_l2_found and not (begin_found["Page"] is None) and (end_found["Page"] is None):
                end_found["Details"] = page.search(pattern=self.markings["end_marking"].format(int(index) + 1))
                if len(end_found["Details"])==1:
                    end_found["Page"]=page.page_number
                    break
        if not (begin_found["Page"] is None) and not (end_found["Page"] is None):
            start=False
            finish=False
            for PID in range(begin_found["Page"],end_found["Page"]+1):
                current_text = self.File.pages[PID-1].extract_text_lines()
                if PID==begin_found["Page"]:
                    for row in current_text:
                        if row['text']==begin_found["Details"][0]["text"]:
                            start=True
                        if start:
                            final_result = final_result + row['text']
                elif PID==end_found["Page"]:
                    for row in current_text:
                        if row['text'] == end_found["Details"][0]["text"]:
                            finish = True
                        if not finish:
                            final_result = final_result + row['text']
                else:
                    for row in current_text:
                        final_result=final_result+row['text']
        return self.Bu,final_result
