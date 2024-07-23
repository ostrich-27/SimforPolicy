import os


class GenerateConfig(object):
    def __init__(self,f):
        self.F=f
        self.__main()

    def __main(self):
        Result="Code|Path|Status\n"
        names = os.listdir(self.F)
        for name in names:
            name_arr = name.split(".")
            if name_arr[len(name_arr) - 1] == "pdf" or name_arr[len(name_arr) - 1] == "PDF":
                code=name.replace(".pdf","")
                code=code.replace(".PDF","")
                Path = os.path.join(self.F, name)
                Result=Result+"{}|{}|{}\n".format(code,Path,"")
        with open(r"Config.txt","w+") as f:
            f.write(Result)

if __name__=="__main__":
    g=GenerateConfig(r"C:\Users\MY598RL\Desktop\pdf_files_2023_ALL")
