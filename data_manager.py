import os
import docx
import time
import openai
from zipfile import ZipFile
from bs4 import BeautifulSoup
import collections
import pandas as pd
import json
import itertools
import numpy as np
import streamlit as st
from streamlit_echarts import st_pyecharts
from pyecharts import options as opts
from pyecharts.charts import Graph
from pyecharts.globals import ThemeType
from difflib import get_close_matches
from sentence_transformers import SentenceTransformer,util

prompt_templete  = ""

class DataBase:
    def __init__(self,path):
        self.path = path
        with open(path,"r",encoding="utf-8") as f:
            self.data = json.load(f)
        self.model = None
        
    
    def save(self):
        with open(self.path,"w",encoding='utf-8') as f:
            json.dump(self.data,f,indent=4,ensure_ascii=False)

    def get_api_result(self,prompt,model="gpt-3.5-turbo-0613"):
        try:
            completion = openai.ChatCompletion.create(
                model=model,
                messages=[
                    {"role": "system", "content": "You are a powerful assistant."},
                    {"role": "user", "content": prompt}],
                temperature=0
                )
            answer = completion["choices"][0]['message']['content']
            return (True,answer)
        except Exception as error:
            print(f"API ERROR: {error}")
            if "Rate limit" in str(error):
                time.sleep(2)
            if "empty message" in str(error):
                print(prompt)
            if "maximum context length" in str(error):
                return self.get_api_result(prompt,model="gpt-3.5-turbo-16k-0613")
            return (False,"Error")

    def get_author_keywords(self,paras):
        text = " ".join(paras)

        # 调用api
        prompt = prompt_templete.format(input=text)
        error_num = 0
        while True:
            status,api_output = self.get_api_result(prompt)
            if api_output == '':
                print("Api output is none,prompt:",prompt)
                status = False
            if status:
                break
            else:
                time.sleep(1)
                error_num += 1
                # 重复次数过多
                if error_num == 5:
                    print("重复次数过多！")
                    time.sleep(10)
                    break
        
        # 解析api结果
        try:
            api_output = eval(api_output)
            keywords = api_output['keyword']
            author = api_output['author']
            return author,keywords
        except:
            return [],[]
        

    def add_paper(self,filepath):
        """
        添加新论文流程:
            1. 检测是否与数据库中的存在重复 
            2. 将其赋上id以及年份
            3. 调用api得到其作者、关键词
        """
        if "docx" not in filepath:
            return False
        # 标题
        title1 = item['filename'].replace("docx","")
        title2 = item['text'][0]
        if len(title2) < len(title1)/2:
            title = title1
        else:
            title = title2
        item['title'] = title

        # 年份
        year = time.ctime(os.path.getmtime(filepath)).split(" ")[-1]
        # id: 当前最大id+1
        max_id = max([item['id'] for item in self.data])
        _,filename = os.path.split(filepath)
        item = {"filename":filename,
                "id":max_id+1,
                "year":year}
        # 读取文档内容
        try:
            file = docx.Document(filepath)
            paras = [i.text for i in file.paragraphs if i.text != ""]
        except Exception as e:
            document = ZipFile(filepath)
            xml = document.read("word/document.xml")
            wordObj = BeautifulSoup(xml.decode("utf-8"))
            texts = wordObj.findAll("w:t")
            paras = [text.text for text in texts]
        if paras == []:
            return 
        item['text'] = paras

        # 调用LLM获取关键词及作者名称
        author,keywords = self.get_author_keywords(paras)
        item['author'] = author
        item['keywords'] = keywords

    def build_excel(self,save_path = "output.xlsx"):
        have_author_paper = []
        for item in self.data:
            if item['author'] != []:
                have_author_paper.append(item)
        all_author = list(set(sum([item['author'] for item in have_author_paper],[])))
        author_dict = collections.defaultdict(list)
        for author in all_author:
            for paper in have_author_paper:
                if paper['text'] == []:
                    continue
                if author in paper['author']:
                    # 作者顺序
                    order = paper['author'].index(author)+1
                    keywords = "、".join(paper['keywords'])
                    year = paper['year']
                    title = paper['filename'].replace("docx","")
                    co_authors = [i for i in paper['author'] if i!=author]
                    co_authors = "、".join(co_authors)
                    author_dict[author].append([title,order,year,keywords,co_authors])
        output = []
        map_dict = {1:"一",2:"二",3:"三",4:"四",5:"五",6:"六"}
        for author,papers in author_dict.items():
            for paper in papers:
                title,order,year,keywords,co_authors = tuple(paper)
                output.append([author,title,map_dict[order],year,keywords,co_authors])
        output = pd.DataFrame(output,columns=['专家名',"文章标题","作者排序","年份","专业标题（关键词）","合作者"])
        # 专家人名只保留一个
        last_name = ""
        for i in range(output.shape[0]):
            name = output.iloc[i,0]
            if name == last_name:
                output.iloc[i,0] = ""
            else:
                last_name = name
        output.to_excel(save_path,index=False)

    def build_graph(self):
        new_items = []
        for item in self.data:
            if "embedding" in item:
                new_items.append(item)
        nodes = []
        links = []
        categories = []

        # 提取所有作者
        all_author = list(set(sum([item['author'] for item in new_items],[])))
        author_dict = collections.defaultdict(list)
        for author in all_author:
            for paper in new_items:
                if paper['text'] == []:
                    continue
                if author in paper['author']:
                    # 作者顺序
                    order = paper['author'].index(author)+1
                    keywords = "、".join(paper['keywords'])
                    year = paper['year']
                    title = paper['title']
                    co_authors = [i for i in paper['author'] if i!=author]
                    co_authors = "、".join(co_authors)
                    author_dict[author].append([title,order,year,keywords,co_authors])

        categories = [{"name":i} for i in all_author] 
        nodes = [{"name":key,
                "symbolSize":len(value)*10 + 30,
                "value":len(value),
                'category':key,
                "label": {"normal":{"show":"True"}}} for key,value in author_dict.items()] + [{"name":item['title'],
                                                                                                "symbolSize":10,
                                                                                                "value":1,
                                                                                                'category':item['author'][0],
                                                                                                'itemStyle':{"normal":{"color":"#2f4554"}}} for item in new_items]

        co_author_dict = {}
        # 共同作者进行连线
        for item in new_items:                                                                   
            for each in itertools.combinations(item['author'], 2):
                if each != []:
                    a1,a2 = tuple(sorted(list(each)))
                    if (a1,a2) not in co_author_dict:
                        co_author_dict[(a1,a2)] = 1
                    else:
                        co_author_dict[(a1,a2)] += 1
        for key,value in co_author_dict.items():
            links.append({"source":key[0],"target":key[1],"value":value*10})

        # 作者与文章连线
        for item in new_items:
            for author in item['author']:
                links.append({"source":author,"target":item['title'],"value":100})

        # 相似度高的文章连线
        docs = [item['title'] for item in new_items]
        embeddings = [item['embedding'] for item in new_items]
        embeddings = np.array(embeddings)
        norm = np.linalg.norm(embeddings,axis=-1,keepdims=True)
        arr_norm = embeddings / norm
        cosine_matrix = np.dot(arr_norm,arr_norm.T)
        for i,row in enumerate(cosine_matrix):
            for j,sim in enumerate(row):
                if i <= j:
                    continue
                if sim > 0.5:
                    links.append({"source":docs[i],"target":docs[j],"value":10*sim})

        c = (
            Graph(init_opts=opts.InitOpts(theme=ThemeType.PURPLE_PASSION,
                                          width="1000px", height="600px"))
            .add(
                "",
                nodes,
                links,
                categories,
                repulsion=1000,
                linestyle_opts=opts.LineStyleOpts(curve=0.4,color="source"),
                label_opts=opts.LabelOpts(is_show=False),
            )
            .set_global_opts(
                legend_opts=opts.LegendOpts(is_show=False),
                title_opts=opts.TitleOpts(title="论文关系网络"),
            )
            #.render("graph_weibo.html")
        )
        return c

    def search_name(self,query):
        map_dict = {1:"一",2:"二",3:"三",4:"四",5:"五",6:"六"}
        # 提取所有作者
        all_author = list(set(sum([item['author'] for item in self.data],[])))
        author_dict = collections.defaultdict(list)
        for author in all_author:
            for paper in self.data:
                if paper['text'] == []:
                    continue
                if author in paper['author']:
                    # 作者顺序
                    order = paper['author'].index(author)+1
                    order = map_dict[order]
                    keywords = "、".join(paper['keywords'])
                    year = paper['year']
                    title = paper['title']
                    co_authors = [i for i in paper['author'] if i!=author]
                    co_authors = "、".join(co_authors)
                    author_dict[author].append([title,order,year,keywords,co_authors])

        # 先尝试找完全匹配的人名
        matches = [name for name in all_author if query.lower() == name.lower()]
        if not matches:
            matches = get_close_matches(query, all_author, n=5, cutoff=0.6)
        detail = []
        matches = list(set(matches))
        for name in matches:
            detail.append((name,author_dict[name]))
        return detail

    def search_keywords(self,query):
        if self.model == None:
            self.model = SentenceTransformer(r'E:\code\qe\bert')
            self.model.max_seq_length = 64
        query_emb = self.model.encode(query, convert_to_tensor=True,batch_size=1)
        new_items = []
        for item in self.data:
            if "embedding" in item:
                new_items.append(item)
        embeddings = [item['embedding'] for item in new_items]
        doc_emb = np.array(embeddings,dtype=np.float32)
        cosine_scores = util.cos_sim(query_emb,doc_emb)[0,:].numpy()
        top10_index = cosine_scores.argsort()[::-1][:10]
        top10_item = []
        for index in top10_index:
            top10_item.append(new_items[index])
        return top10_item

database = DataBase("database_new.json")
graph = database.build_graph()
# 设置Streamlit的页面标题
st.title("论文检索系统")

# 在侧边栏添加搜索的组件
selected_option = st.sidebar.selectbox("检索方式", ["按作者", "按关键词"])
search_query = st.sidebar.text_input("查询内容", "")
submit_button = st.sidebar.button("查询")


if submit_button:
    st.markdown("#### 检索结果")
    if selected_option == "按作者":
        result = database.search_name(query=search_query)
        if result == []:
            st.write("未查找到该作者!")
        else:
            for item in result:
                author,detail = item
                st.markdown(f"#### {author}")
                for index,item in enumerate(detail):
                    st.markdown(f"##### {index+1}. {item[0]}")
                    st.write(f"关键词: {item[-2]}")
                    st.write(f"年份: {item[2]}")
                    co_author = "无" if item[-1] == "" else item[-1]
                    st.write(f"第{item[1]}作者, 合作者: {co_author}")
    else:
        result = database.search_keywords(query=search_query)
        for index,item in enumerate(result):
            st.markdown(f"#### {index+1}. {item['title']}")
            st.write(f"{'、'.join(item['author'])}, {item['year']}")
            st.write(f"关键词: {'、'.join(item['keywords'])}")

st.markdown("---")
# 将Pyecharts图表渲染为图片并显示在页面上
with st.container():
    st_pyecharts(graph,width="1000px", height="600px")






       
        


    