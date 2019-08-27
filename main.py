import sys
from glob import glob
import itertools

import MeCab
import numpy as np
from sklearn.feature_extraction.text import CountVectorizer

from sub.excel1 import excel2doc
from sub.pdf1 import pdf2doc

def cos_sim(v1, v2):
    return np.dot(v1, v2) / (np.linalg.norm(v1) * np.linalg.norm(v2))


if __name__ == "__main__":
    # 検索ディレクトリを入力
    path = input("検索するディレクトリ:")
    # ファイルパス
    excel_path_list = glob(path + "/*.xlsx")
    pdf_path_list = glob(path+"/*.pdf")

    # ファイルがない場合は終了
    if len(excel_path_list) + len(pdf_path_list) == 0:
        print("ファイルがありませんでした")
        sys.exit()

    # ファイル名から文書を抜き出し
    excel_data = {name: excel2doc(name) for name in excel_path_list}
    pdf_data = {name: pdf2doc(name) for name in pdf_path_list}

    # 分かち書き

    corpus = []
    col_dict={}
    i = 0
    for key, item in excel_data.items():
        corpus.append(item)
        col_dict[i] = key
        i += 1

    for key, item in pdf_data.items():
        corpus.append(item)
        col_dict[i] = key
        i += 1

    tagger = MeCab.Tagger('-Owakati')
    corpus = [tagger.parse(sentence).strip() for sentence in corpus]

    vectorizer = CountVectorizer(token_pattern=u'(?u)\\b\\w+\\b')
    bag = vectorizer.fit_transform(corpus)
    bag_of_words = bag.toarray()

    #組み合わせを考える
    for col in range(len(col_dict)-1):
        for row in range(1+col, len(col_dict)):
            #類似度
            similar = cos_sim(bag_of_words[col], bag_of_words[row])

            if similar >=0.7 :
                print("{}と{}は類似しています。".format(col_dict[col], col_dict[row]))


