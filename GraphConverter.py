#Graph Converter - v1.0.0
import networkx as nx
import pandas as pd
import xlsxwriter 
import warnings
import pickle
import sys
import os

#GUI Imports:
from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter.filedialog import asksaveasfile 
from tkinter.filedialog import askdirectory
from tkinter import messagebox

##Global variables##
G = None
excelMatrix=None
directory = "."
saveFolder = "."
waitForMatrix = True
matrixName = ""
path=''

#xlsxWriter
workbook = None
worksheet = None
#xlsx Formats
main_format=None
side_format=None
mixed_format=None

#Matrix Data
df = None

#Utility Functions
def isSquare (m): return all (len (row) == len (m) for row in m)
def OpenFile(name): sys.stdout = open(name, "w")
def CloseFile(): sys.stdout.close()
def isDirected():
    global G
    directed = False
    for x,y in G.edges():
        if G.get_edge_data(x,y)!=G.get_edge_data(y,x): 
            directed = True        
    return directed

#Input Functions
def OpenMatrix(filepath):
    excelMatrix = pd.read_excel(filepath, sheet_name=0, header=0,index_col=0)
    return excelMatrix
def read_dl(filename):
    header=0
    try:            
        with open(filename) as myFile:
            for num, line in enumerate(myFile, 1):
                if ("data:") in line.lower():
                    header=num
        dl=pd.read_csv(filename,sep=" ",header=None,skiprows=header)
    except pd.errors.ParserError as e:
        return None    
    return dl    
def read_pickle(filename):
   # open a file, where you stored the pickled data
   G = pickle.load(open(filename, 'rb'))
   return G

#Output Functions
def write_dl(G,filename):
    G=G.to_directed()
    isW=nx.is_weighted(G)
    if isW:
        isReallyW=False
        for x,y in G.edges():
            if G.get_edge_data(x,y)['weight']!=1: 
                isReallyW=True
                break
        if isReallyW==False:
            isW=False

    with open(filename,'w') as f:
        f.write("DL n=" + str(len(G))+"\n")
        f.write("format = edgelist1\n")
        f.write("labels embedded:\n")
        f.write("data:\n\n")
        
        if isW:
            for line in nx.generate_edgelist(G,data=['weight'] ):
                f.write(str(line))
                f.write("\n")
        else:        
            for line in nx.generate_edgelist(G):
                f.write(str(line.split(' {')[0]))
                f.write("\n")
        for node in G.nodes():
            if G.degree[node]==0:
                f.write(str(node))
                f.write("\n")
def write_pickle(G,filename):
    pickle.dump(G, open(filename, 'wb'))

def write_xlsx(G,filename):
    table=nx.to_pandas_adjacency(G,nodelist=sorted(G.nodes),weight='weight',nonedge=float("NaN"))

    excelSaved=False
    while not excelSaved:
        try:
            table.to_excel(filename)
            excelSaved=True  
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo("Output error", "Excel File already open! Close it and click OK")    
def write_edgelist(G,filename):
    #XlsxWriter Start
    global worksheet
    global main_format
    global side_format
    global mixed_format

    G=G.to_directed()
    isW=nx.is_weighted(G)

    if isW:
        isReallyW=False
        for x,y in G.edges():
            if G.get_edge_data(x,y)['weight']!=1: 
                isReallyW=True
                break
        if isReallyW==False:
            isW=False

    workbook = xlsxwriter.Workbook(filename, {'nan_inf_to_errors': True})  
    worksheet = workbook.add_worksheet("Graph")   

    main_format = workbook.add_format()
    main_format.set_bold()
    main_format.set_border(2)
    side_format = workbook.add_format()    
    side_format.set_border(1)
    mixed_format = workbook.add_format()    
    mixed_format.set_bold()  
    mixed_format.set_border(1)

    #Formatting
    worksheet.set_column(2, 3, 20) 

    #Excel code    
    worksheet.write(0,0,"Source", main_format)  
    worksheet.write(0,1,"Target", main_format) 
    if isW:
        worksheet.write(0,2,"Weight", main_format)  

    lineCounter=1
    for line in nx.generate_edgelist(G,delimiter=';!%',data=['weight'] ):
        edge=line.split(';!%')
        worksheet.write(lineCounter,0,edge[0], side_format)  
        worksheet.write(lineCounter,1,edge[1], side_format)
        if isW:
            worksheet.write(lineCounter,2,float(edge[2]), side_format)  
        lineCounter+=1
    for node in G.nodes():
        if G.degree[node]==0:
            worksheet.write(lineCounter,0,str(node), side_format)  
            lineCounter+=1

    #XlsxWriter End
    excelSaved=False
    while not excelSaved:
        try:
            workbook.close()
            excelSaved=True  
        except xlsxwriter.exceptions.FileCreateError:
            messagebox.showinfo("Output error", "Excel File already open! Close it and click OK")

#Converting Function
def ConvertClick():
    global waitForMatrix
    global graphName
    global G
    global path    

    if waitForMatrix:
        messagebox.showinfo("Alert", "Select an input graph first!")
    else:  
        analyzeBtn.config(state=DISABLED) 
        analyzeBtnText.set("Converting...")          
        root.update()   

        #Graph Creation code
        excelMatrix=None
        G=None
        usingEdgeList=False

        if matrixName.endswith(".xlsx") or matrixName.endswith(".xls"):    
            excelMatrix=OpenMatrix(path)
            if isSquare(excelMatrix.to_numpy()): 
                excelMatrix.fillna(0, inplace=True)
                G = nx.from_pandas_adjacency(excelMatrix, create_using=nx.DiGraph()) 
            else:
                excelMatrix.reset_index(inplace=True)
                try:
                    excelMatrix.columns = ['from','to','weight',]
                except ValueError:
                    try:
                        excelMatrix.columns = ['from','to']
                        excelMatrix['weight'] = 1
                    except ValueError:
                        messagebox.showinfo("Input error", "Invalid Excel input! Please double-check!")
                        return
                usingEdgeList=True    
        elif matrixName.endswith(".nxg"):
            G=read_pickle(path)
        else:        
            excelMatrix=read_dl(path) 
            if excelMatrix is None:
                return
            try:
                excelMatrix.columns = ['from','to','weight']
            except ValueError:
                try:
                    excelMatrix.columns = ['from','to']
                    excelMatrix['weight'] = 1
                except ValueError:
                    messagebox.showinfo("Input error", "Invalid DL input! Please double-check!")
                    return
                
            usingEdgeList=True

        if usingEdgeList:            
            try:
                G=nx.from_pandas_edgelist(excelMatrix, 'from','to','weight', create_using=nx.DiGraph())
            except ValueError:
                messagebox.showinfo("Input error", "Input isn't a matrix or an edgelist!")
                return
            edgesToRemove = []     
            nodesToRemove = []
            for x,y in G.edges():
                if pd.isnull(y):
                    edgesToRemove.append([x,y])                    
                    nodesToRemove.append(y)
                G.add_node(x)             
            for z in edgesToRemove:
                G.remove_edge(z[0], z[1])
            for z in nodesToRemove:
                try:
                    G.remove_node(z) 
                except nx.exception.NetworkXError:
                    pass   

        if isDirected():
            isD=True
            geo = nx.DiGraph() 
            for x in sorted(G.nodes()):
                geo.add_node(x)
            for x in sorted(G.edges(data=True)):
                geo.add_edge(x[0],x[1],weight=x[2]['weight'])
            G=geo
        else:
            isD=False
            geo = nx.Graph() 
            for x in sorted(G.nodes()):
                geo.add_node(x)
            for x in sorted(G.edges(data=True)):
                geo.add_edge(x[0],x[1],weight=x[2]['weight'])
            G=geo                 
        #Graph Creation End
        outputName=""
        if graphsOutput.get()=="DL":
            outputName=str(graphName) + " DL.txt"
            write_dl(G, str(graphName) + " DL.txt")
        if graphsOutput.get()=="Matrix":
            if len(G.nodes())>16000:
                messagebox.showinfo("Alert", "Graph too big for excel!")
            else:
                outputName=str(graphName) + " Matrix.xlsx"
                write_xlsx(G, str(graphName) + " Matrix.xlsx") 
        if graphsOutput.get()=="Edge List":
            outputName=str(graphName) + " Edge List.xlsx"
            write_edgelist(G, str(graphName) + " Edge List.xlsx") 
        if graphsOutput.get()=="NetworkX Graph":
            outputName=str(graphName) + " Pickle.nxg"
            write_pickle(G, str(graphName) + " Pickle.nxg") 

        G=None
        excelMatrix=None
        analyzeBtn.config(state=DISABLED)  
        analyzeBtnText.set("All done!")            

#GUI code
def OpenFileClick():
    root.update() 
    global directory
    global graphName
    global path
    global matrixName
    global waitForMatrix

    path = askopenfilename(initialdir = directory, title = "Select the input graph:", filetypes =[("Graph input files", "*.xlsx *.xls *.txt *.nxg")])
    directory = os.path.dirname(path) 
    matrixName = os.path.basename(path)  
    graphName=matrixName.rsplit('.', 1)[0]

    if graphName!='':
        text.set(matrixName)
        waitForMatrix=False
        directory = os.path.dirname(path)    
        analyzeBtnText.set("Convert")
        analyzeBtn.config(state=NORMAL)         
def ClearProgress(*something): 
    analyzeBtnText.set("Convert")
    if waitForMatrix == False:
        analyzeBtn.config(state=NORMAL) 
def center(win):
    win.update_idletasks()
    width = win.winfo_width()
    frm_width = win.winfo_rootx() - win.winfo_x()
    win_width = width + 2 * frm_width
    height = win.winfo_height()
    titlebar_height = win.winfo_rooty() - win.winfo_y()
    win_height = height + titlebar_height + frm_width
    x = win.winfo_screenwidth() // 2 - win_width // 2
    y = win.winfo_screenheight() // 2 - win_height // 2
    win.geometry('{}x{}+{}+{}'.format(width, height, x, y))
    win.deiconify()

root = Tk()
root.geometry("290x120") 
root.resizable(0, 0)
center(root)
root.title("Graph Converter v1.0.0")

tlabel = Label(root, text ="Input Graph:")
tlabel.grid(row=0, column=0, padx = 10, pady = (10,0), sticky=W)

text=StringVar()
text.set("Select the input graph...")
Btn = Button(root, textvariable=text, command = OpenFileClick, width=22).grid(row=0, column=1, padx = 10,pady = (10,0), sticky=W)

labelGraphsOutput = Label(root, text ="Output Format:")
labelGraphsOutput.grid(row=1, column=0,  padx=10, pady=5, sticky=W)

graphsOutput = StringVar()
graphsOutput.set("Matrix")
graphsOutputMenu = OptionMenu(root, graphsOutput, "DL", "Matrix", "Edge List", "NetworkX Graph",command=ClearProgress)
graphsOutputMenu.grid(row=1,column=1,columnspan=2, padx=8, pady=5, sticky=EW)

analyzeBtnText = StringVar()
analyzeBtnText.set("Convert")
analyzeBtn = Button(root, textvariable=analyzeBtnText, command = ConvertClick, width=20)
analyzeBtn.grid(row=10,  column=0, padx=10, pady=10, columnspan=5, sticky=EW)
analyzeBtn.config(state=DISABLED)    

root.mainloop()
