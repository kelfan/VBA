#!/usr/bin/python
# -*- coding: UTF-8 -*-
import sklearn.datasets as datasets
import pandas as pd
import os
import sys, getopt

def main(argv):
    path = ''
    try:
        opts, args = getopt.getopt(argv,"hp:",["path="])
    except getopt.GetoptError:
        print ('test.py -p <path>')
        sys.exit(2)
    for opt, arg in opts:
        if opt == '-h':
            print ('test.py -p <path>')
            sys.exit()
        elif opt in ("-p", "--path"):
            path = arg
    # print ('输入的文件为：', path + '\dt.png')
    
    iris=datasets.load_iris()
    df=pd.DataFrame(iris.data, columns=iris.feature_names)
    y=iris.target
    from sklearn.tree import DecisionTreeClassifier
    dtree=DecisionTreeClassifier()
    dtree.fit(df,y)
    from sklearn.externals.six import StringIO  
    from IPython.display import Image  
    from sklearn.tree import export_graphviz
    import pydotplus
    dot_data = StringIO()
    export_graphviz(dtree, out_file=dot_data,filled=True, rounded=True,special_characters=True)
    graph = pydotplus.graph_from_dot_data(dot_data.getvalue())  
    Image(graph.create_png())
    graph.write_png(path + '\dtTest.png')
    print(os.getcwd())

if __name__ == "__main__":
    main(sys.argv[1:])