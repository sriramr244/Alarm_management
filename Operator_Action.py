import pandas as pd
import numpy as np
import os
from os import walk
import re
import sys
import getpass
import time
import matplotlib
from pylab import title, xlabel, ylabel, xticks, bar, legend, axis, savefig, figure
from fpdf import FPDF
import seaborn as sns
import random
import csv
from progress.bar import Bar


''' Update notes --- Beta V.1.1
Added user more and configuration mode. Configuration mode allows us to reset the dictionary files
Added config file to be able to work with
Pending work:
Change the code for resume logic to be able to handle resume from the start. 
'''
#Function that will import all excel files in a folder
#It will take in the path as input and return a dataframe will contain a concatination of all entries of all three files

def open_mes_map():
    #Replace the path here with settings dict path
    with open(os.getcwd()+"\\Utility_Files\\"+"MES_Path_Map.csv", mode='r') as infile:
        reader = csv.reader(infile)
        mes_map = {rows[0]:rows[1] for rows in reader}
    return mes_map

def norm(path):
    return path.split("]")[1].split("/",2)[2]

def convert_to_ticks(value):
    value = value.replace(" ","-").replace(":","-").replace(".","-").split("-")
    return time.mktime((int(value[0]), int(value[1]), int(value[2]), int(value[3]), int(value[4]), int(value[5]), 0 , 0, 0))

def normalize_time_mes(df):
    start_time = []
    end_time = []
    for i in range(len(list(df.loc[:,"Start Time"]))):
        start_time.append(convert_to_ticks(str(df.loc[i,"Start Time"])))
        end_time.append(convert_to_ticks(str(df.loc[i,"End Time"])))
    df["Start_Time_Ticks"] = start_time
    df["End_Time_Ticks"] = end_time
    return df.sort_values('Start_Time_Ticks',ascending = True).replace(np.nan,"nan",regex = True)

def is_in(key_words,chunk):
    for word in key_words:
        if word.lower() in chunk.lower():
            return True
        else:
            continue
    return False

def map_from_mes_data(df,mes_df, mes_map):
    po = []
    batch = []
    source = []
    material = []
    phase_no = []
    setpoint = []
    dest = []
    line = []
    main_equip = []
    main_eq_match = []
    print("Mapping MES Data...")
    bar = Bar("Data Mapping", max=len(list(df.loc[:,"Path"].values)))
    for i in range(len(list(df.loc[:,"Path"].values))):
        #if(i%1000 == 0):
            #print(str(i),'/',str(len(list(df.loc[:,"Path"].values))),' records processed', end='\r')
        bar.next()
        timetick = df.loc[i,"TimeTicks"]
        if df.loc[i,"Object_Type"] not in settings_dict["MES_report_obj_to_incl"].split(","): #changed
            po.append('nan')
            batch.append('nan')
            source.append('nan')
            material.append('nan')
            phase_no.append('nan')
            setpoint.append('nan')
            dest.append('nan')
            line.append('nan')
            main_equip.append('nan')
            main_eq_match.append('nan')
            continue
        try:
            equipment = mes_map[norm(df.loc[i,"Path"])]
        except:
            po.append('nan')
            batch.append('nan')
            source.append('nan')
            material.append('nan')
            phase_no.append('nan')
            setpoint.append('nan')
            dest.append('nan')
            line.append('nan')
            main_equip.append('nan')
            main_eq_match.append('nan')
            continue
        ld_check = False
        me_check = False
        mes_df_2 = mes_df[(mes_df["Start_Time_Ticks"]<= timetick) & (mes_df["End_Time_Ticks"]>= timetick)]
        if is_in(settings_dict["Tra_keys"].split(","), df.loc[i,"Path"]):
            mes_df_2=mes_df_2[(mes_df_2["Destination"].str.contains(equipment))]
            ld_check = True
        elif is_in(settings_dict["Line_keys"].split(","), equipment):
            mes_df_2=mes_df_2[(mes_df_2["Line"].str.contains(equipment))]
            ld_check = True
        else:
            pass
        if (ld_check==False) or (len(list(mes_df_2.loc[:,"Phase No"]))<1):
            mes_df_2=mes_df_2[(mes_df_2["Main Equipment"].str.contains(equipment))]
            me_check=True
        mes_df_2 = mes_df_2.reset_index()
        if len(list(mes_df_2.loc[:,"Phase No"]))>0:
            if not(me_check):
                po.append(mes_df_2.loc[0,"PONO"])
                batch.append(mes_df_2.loc[0,"Product Name"])
                source.append(mes_df_2.loc[0,"Source"])
                material.append(mes_df_2.loc[0,"Material Name"])
                phase_no.append(mes_df_2.loc[0,"Phase No"])
                setpoint.append(mes_df_2.loc[0,"Std. Quantity"])
                dest.append(mes_df_2.loc[0,"Destination"])
                line.append(mes_df_2.loc[0,"Line"])
                main_equip.append(mes_df_2.loc[0,"Main Equipment"])
                main_eq_match.append('No')
            else:
                po.append(mes_df_2.loc[0,"PONO"])
                batch.append(mes_df_2.loc[0,"Product Name"])
                source.append(mes_df_2.loc[0,"Source"])
                material.append(mes_df_2.loc[0,"Material Name"])
                phase_no.append(mes_df_2.loc[0,"Phase No"])
                setpoint.append(mes_df_2.loc[0,"Std. Quantity"])
                dest.append(mes_df_2.loc[0,"Destination"])
                line.append(mes_df_2.loc[0,"Line"])
                main_equip.append(mes_df_2.loc[0,"Main Equipment"])     
                main_eq_match.append('Yes')     
        else:
            po.append('nan')
            batch.append('nan')
            source.append('nan')
            material.append('nan')
            phase_no.append('nan')
            setpoint.append('nan')
            dest.append('nan')
            line.append('nan')
            main_equip.append('nan')
            main_eq_match.append('nan')
    df["PO"]=po
    df["Batch_Material"]=batch
    df["Source"] = source
    df["Material"]=material
    df["Phase No"]=phase_no
    df["Set_Point"]=setpoint
    df["Destination"]=dest
    df["line"]=line
    df["Main Equipment"]=main_equip
    df["Mapped with Main Eq"]=main_eq_match
    bar.finish()
    print("")
    print(str(len(list(df.loc[:,"Path"].values))),'/',str(len(list(df.loc[:,"Path"].values))),' records processed', end='\n')
    print("MES Data Mapped Successfully")

def convert_df_to_plot(df_input):
    df_input = df_input.head(int(settings_dict["Include_top"]))
    col = list(df_input.columns)
    pv = []
    count = []
    duration = []
    for i in range(len(col)-1):
        pv.append(list(df_input.loc[:,col[0]]))
        for j in range(6):
            count.append(df_input.iloc[j,i+1])
            duration.append(col[i+1])
    pv = sum(pv,[])
    df_plot = pd.DataFrame()
    df_plot[col[0]] = pv
    df_plot["Count"] = count
    df_plot["Period"] = duration
    return df_plot

def process(value):
    return value.split(' ')[1].split('-')[0]
def process1(value):
    return value.split(' ')[1].split('-')[1]
def process2(value):
    if int(value.split(' ')[1].split('-')[1])<=15:
        return 'FN1'
    return 'FN2'

def pivott_v3(df, FP, FV, PE, FN):
    mon = ['January','February','March', 'April','May','June','July','August','September','October','November','December']
    data=df.copy()
    for i in range(0,len(FP)):
        filt = FP[i]
        Var = FV[i]
        for j in Var:
            data = data[data[filt] == j]
    data['month']=data['EventTime'].apply(process)
    data['date']=data['EventTime'].apply(process1)
    #data = data.groupby(PE).size().rename('Count').to_frame()
    if FN==0:
        MO = list(set(data['month']))
        MO.sort()
        Month = dict()
        for k in MO:
            mOnth = mon[int(k)-1]
            data1 = data[data['month']==k]
            #data.head()
            Result=dict()
            for i in data1[PE]:
                Result[i] = Result.get(i, 0) + 1
            for j in df[PE]:
                if j not in list(Result.keys()):
                    Result[j]=0
            sorted_keys = sorted(Result, key=Result.get, reverse=True)
            sorted_Result=dict()
            for w in sorted_keys:
                sorted_Result[w] = Result[w]
            Month[mOnth]=sorted_Result
        m=0
        for Title, Counts in Month.items():
            if m==0:
                df_ = pd.DataFrame()
                df_[PE]=list(Counts.keys())
                df_[Title]=list(Counts.values())
                m+=1
            lis=[]
            for l in list(df_[PE]):
                lis.append(Counts[l])
            df_[Title]=lis        
        return df_
    if FN==1:
        data['Fo_Ni']=data['EventTime'].apply(process2)
        data[(data['Fo_Ni']=='FN1')&(data['month']=='10')]
        MO = list(set(data['month']))   
        MO.sort()
        FN = dict()
        for k in MO:
            mOnth = mon[int(k)-1]
            data1 = data[data['month']==k]
            FO = list(set(data1['Fo_Ni']))
            FO.sort()
            for l in FO:
                ref = mOnth+' '+l
                data2 = data1[data1['Fo_Ni']==l]
                Result = dict()
                for i in data2[PE]:
                    Result[i] = Result.get(i, 0) + 1
                for j in df[PE]:
                    if j not in list(Result.keys()):
                        Result[j]=0
                sorted_keys = sorted(Result, key=Result.get, reverse=True)
                sorted_Result=dict()
                for w in sorted_keys:
                    sorted_Result[w] = Result[w]
                FN[ref]=sorted_Result
        m=0
        for Title, Counts in FN.items():
            if m==0:
                df_ = pd.DataFrame()
                df_[PE]=list(Counts.keys())
                df_[Title]=list(Counts.values())
                m+=1
            lis=[]
            for l in list(df_[PE]):
                lis.append(Counts[l])
            df_[Title]=lis
        return df_

def process_allowable_interventions(df):
    df = df.reset_index()
    path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"allowed_interventions.txt"
    allowable_indices = []
    #allowable_timeticks = []
    print("Processing allowable interventions...")
    try:
        with open(path) as f:
            filt = f.read()
    except:
        print("Failed to open allowed_interventions.txt... Skipping this step")
        return df
    filt = filt.replace("\n","").replace(" ","").replace("&","=").replace("Acceptable{[","").replace("]}","").replace("%"," ").split(";")
    if len(filt[-1]) == 0:
        filt.pop()
    for line in filt:
        line = line.split('=')
        i=0
        data = df.copy()
        while(i<len(line)):
            data = data[data[line[i]] == line[i+1]]
            i+=2
        allowable_indices=allowable_indices+(list(data.index))
        #allowable_timeticks=allowable_timeticks+(list(data.loc[:,"TimeTicks"].values))
    i=0
    count = 0
    l = len(set(allowable_indices))
    if len(allowable_indices)>0:
        #for entry in list(df.loc[:,"Index"]):
        #    if entry in set(allowable_indices):
        #        df.loc[i,"Action_Class"] = "Allowed"
        #        count+=1
        bar = Bar("Allowable Actions", max=len(set(allowable_indices)))
        for entry in set(allowable_indices):
            bar.next()
            df.loc[entry,"Action_Class"] = "Allowed"
            count+=1
            i+=1
            #if(i%1000 == 0):
            #    print(str(i),'/',str(l),' records processed', end='\r')
    else:
        print("No allowable interventions found")
        return df
    bar.finish()
    print("")
    print(str(i),'/',str(l),' records processed', end='\n')
    print("Successfully processed "+str(count)+" allowed interventions")
    return df

def apply_filter(filter_list,df):
    fltr_cnt = len(filter_list)
    print("Applying ",fltr_cnt," filters...")
    #try:
    for filter in filter_list:
        df_copy = df.copy()
        filter_data(df_copy,filter)
    #except:
     #   print("Error applying filter")
      #  sys.exit("Exiting...")
    print("Filters applied successfully")
    return 0

def pivott(df, FP, FV, PE):
    #Function to filter and pivot on a variable. This is used in report generation
    data=df.copy()
    for i in range(0,len(FP)):
        filt = FP[i]
        Var = FV[i]
        for j in Var:
            data = data[data[filt] == j]
    Result = dict()
    for i in data[PE]:
        Result[i] = Result.get(i, 0) + 1
    sorted_keys = sorted(Result, key=Result.get, reverse=True)
    sorted_Result=dict()
    for w in sorted_keys:
        sorted_Result[w] = Result[w]
    return sorted_Result

def pivott_v2(df, FP, FV, PE):
    data=df.copy()
    for i in range(0,len(FP)):
        filt = FP[i]
        Var = FV[i]
        for j in Var:
            data = data[data[filt] == j]
    Result = dict()
    for i in data[PE]:
        Result[i] = Result.get(i, 0) + 1
    Result_dict = dict()
    Result_dict[PE] = list(Result.keys())
    Result_dict['Count'] = list(Result.values())
    ret_frame = pd.DataFrame(data=Result_dict).sort_values('Count',ascending=False)
    return ret_frame

def generateReport_v3(df, report_path=os.getcwd()):
    path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"report_filters.txt"
    temp_path = os.getcwd()+"\\"+"temp"+"\\"
    try:
        print("Reading report settings...")
        heading_height = int(settings_dict["report_heading_height"])
        value_height = int(settings_dict["report_value_height"])
        column_width_var = int(settings_dict["report_column_width_var"])
        column_width_value = int(settings_dict["report_column_width_value"])
        indent = int(settings_dict["report_indent"])
        font = settings_dict["report_font"]
        #to be added more
        print("Settings loaded successfully")
    except:
        print("Load failed... Going with default settings")
        heading_height = 5
        value_height = 5
        column_width_var = 40
        column_width_value = 20
        indent = 15
        font = 'arial'
    with open(path) as f:
        filt = f.read()
    filt=filt.replace('\n',"").replace(" ","").replace("]}","").replace("%"," ")
    filt=filt.replace(']','').split(";")
    if len(filt[-1])==0:
        filt.pop()
        
    Main_result={}
    iter = 0
    for line in filt:
        filt_p=[]
        filt_v=[]
        line = line.split(',')
        if line[0].split('{')[0]=='Filter':
            lin=line[0].replace('Filter{[','')
            lin=lin.split('&')
            
            for k in lin:
                f_V=[]
                m=k.split('=')
                f_V.append(m[1])
                filt_p.append(m[0])
                filt_v.append(f_V)
            PE=line[1].split('=')[1]
            Heading=line[2].split('=')[1] 
            FN = int(line[3].split('=')[1].replace("}",""))
            Main_result[Heading]=pivott_v3(df, filt_p, filt_v, PE, FN)
            
        elif line[0].split('{')[0]=='Generate':
            report_name=line[0].split('{')[1].replace('Filename=','').replace('}','')
            print('Generating ',report_name,'...')
            pdf = FPDF()
            for Heading,data in Main_result.items():
                if len(data.iloc[:,1].values)<1:
                    print("No result for filter entry '",Heading,"'")
                    continue
                cols = list(data.columns)
                pdf.add_page()
                pdf.set_xy(0, 0)
                pdf.set_font(font, 'B', 12)
                pdf.cell(90, 5, " ", 0, 1, 'C')
                pdf.cell(60)
                pdf.cell(75, heading_height, Heading, 0, 1, 'C')
                pdf.cell(90, heading_height, " ", 0, 1, 'C')
                #pdf.cell(-40)
                pdf.set_font(font, 'B', 8)
                pdf.cell(indent)
                pdf.cell(10, heading_height, 'S.No', 1, 0, 'C')
                pdf.cell(column_width_var, heading_height, data.columns[0], 1, 0, 'C')
                for i in range(len(cols)-1):
                    if i==len(cols)-2:
                        pdf.cell(column_width_value, heading_height, cols[i+1], 1, 1, 'C')
                    else:
                        pdf.cell(column_width_value, heading_height, cols[i+1], 1, 0, 'C')
                #pdf.cell(-90)
                pdf.set_font(font, '', 8)
                keyss= list(data.iloc[:,0])
                for i in range(0, len(keyss)):
                    if i==10:
                        break
                    pdf.cell(indent)
                    pdf.cell(10, value_height, '%s' % (str(i+1)), 1, 0, 'C')
                    pdf.cell(column_width_var, value_height, '%s' % (str(keyss[i])), 1, 0, 'R')
                    for j in range(len(cols)-1):
                        if j==len(cols)-2:
                            pdf.cell(column_width_value, value_height, '%s' % (str(data.iloc[i,j+1])), 1, 1, 'L')
                        else:
                            pdf.cell(column_width_value, value_height, '%s' % (str(data.iloc[i,j+1])), 1, 0, 'L')
                    #pdf.cell(-90)
                sns.set(rc={'figure.figsize':(11,11)})
                figure(figsize=(11,11))
                pal = settings_dict["plot_color_palette"]
                if pal == 'random':
                    palettes = ["Blues_d","rocket","deep","vlag","flare","pastel", "mako","crest","magma","viridis","rocket_r","icefire","Spectral","coolwarm"]
                    pal = palettes[random.randrange(0,len(palettes),1)]
                data = convert_df_to_plot(data)
                if len(set(list(data.loc[:,"Period"].values))) > 1:
                    sns_plot = sns.barplot(x = data.columns[0], y = 'Count', data=data, palette=pal, hue = "Period")
                else:
                    sns_plot = sns.barplot(x = data.columns[0], y = 'Count', data=data, palette=pal)
                for p in sns_plot.patches:
                    sns_plot.annotate(format(p.get_height(), '.0f'), 
                                (p.get_x() + p.get_width() / 2., p.get_height()), 
                                ha = 'center', va = 'center', 
                                xytext = (0, 9), 
                                textcoords = 'offset points')    
                #fig = sns_plot.get_figure()
                image_name = 'chart_'+data.columns[0]+report_name+str(iter)+'.png'
                iter+=1
                savefig(temp_path+image_name)
                pdf.cell(90, 1, " ", 0, 2, 'C')
                pdf.image(temp_path+image_name, x = 20, y = None, w = 150, h = 120, type = '', link = '')
                os.remove(temp_path+image_name)
            try:    
                pdf.output(report_path+"\\"+report_name+".pdf", 'F')
                Main_result = {}
            except:
                success = False
                for x in range(11):
                    try:
                        pdf.output(report_name+"_"+str(x+1)+".pdf", 'F')
                        success = True
                        break
                    except:
                        continue
                if not(success):
                    print("Error creating PDF report even after 10 attempts.")
                    return -1
        else:
            print('Invalid Syntax in report_filters.txt... PDF report not generated')
            return -1
            
def generateReport_v2(df):
    path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"report_filters.txt"
    temp_path = os.getcwd()+"\\"+"temp"+"\\"
    with open(path) as f:
        filt = f.read()
    filt=filt.replace('\n',"").replace(" ","").replace("]}","").replace("%"," ")
    filt=filt.replace(']','').split(";")
    if len(filt[-1])==0:
        filt.pop()
        
    Main_result={}
    iter = 0
    for line in filt:
        filt_p=[]
        filt_v=[]
        line = line.split(',')
        if line[0].split('{')[0]=='Filter':
            lin=line[0].replace('Filter{[','')
            lin=lin.split('&')
            
            for k in lin:
                f_V=[]
                m=k.split('=')
                f_V.append(m[1])
                filt_p.append(m[0])
                filt_v.append(f_V)
            PE=line[1].split('=')[1]
            Heading=line[2].split('=')[1].replace('}','')
            Main_result[Heading]=pivott_v2(df, filt_p, filt_v, PE)
            
        elif line[0].split('{')[0]=='Generate':
            report_name=line[0].split('{')[1].replace('Filename=','').replace('}','')+'.pdf'
            print('Generating ',report_name,'...')
            pdf = FPDF()
            for Heading,data in Main_result.items():
                pdf.add_page()
                pdf.set_xy(0, 0)
                pdf.set_font('arial', 'B', 14)
                pdf.cell(90, 10, " ", 0, 1, 'C')
                pdf.cell(60)
                pdf.cell(75, 10, Heading, 0, 1, 'C')
                pdf.cell(90, 10, " ", 0, 1, 'C')
                #pdf.cell(-40)
                pdf.set_font('arial', 'B', 12)
                pdf.cell(40)
                pdf.cell(15, 10, 'S.No', 1, 0, 'C')
                pdf.cell(50, 10, data.columns[0], 1, 0, 'C')
                pdf.cell(30, 10, 'Counts', 1, 1, 'C')
                #pdf.cell(-90)
                pdf.set_font('arial', '', 8)
                keyss= list(data.iloc[:,0])
                values = list(data.iloc[:,1])
                for i in range(0, len(keyss)):
                    if i==10:
                        break
                    pdf.cell(40)
                    pdf.cell(15, 5, '%s' % (str(i)), 1, 0, 'C')
                    pdf.cell(50, 5, '%s' % (str(keyss[i])), 1, 0, 'R')
                    pdf.cell(30, 5, '%s' % (str(values[i])), 1, 1, 'L')
                    #pdf.cell(-90)
                sns.set(rc={'figure.figsize':(11,11)})
                figure(figsize=(11,11))
                sns_plot = sns.barplot(x = data.columns[0], y = 'Count', data=data.head(6), palette="Blues_d")
                for p in sns_plot.patches:
                    sns_plot.annotate(format(p.get_height(), '.0f'), 
                                (p.get_x() + p.get_width() / 2., p.get_height()), 
                                ha = 'center', va = 'center', 
                                xytext = (0, 9), 
                                textcoords = 'offset points')    
                #fig = sns_plot.get_figure()
                image_name = 'chart_'+data.columns[0]+str(iter)+'.png'
                iter+=1
                savefig(temp_path+image_name)
                pdf.cell(90, 10, " ", 0, 2, 'C')
                pdf.image(temp_path+image_name, x = 20, y = None, w = 150, h = 120, type = '', link = '')
                os.remove(temp_path+image_name)
            try:    
                pdf.output(report_name, 'F')
            except:
                print('Error while saving PDF file... Failed to generate PDF report')
                return -1
        else:
            print('Invalid Syntax in report_filters.txt... PDF report not generated')
            return -1

def uniq_id_index(df):
    #input the sorted data frame to add unique id and index
    unique_id = []
    idx = []
    for i in range(len(list(df.loc[:,"ObjectName"].values))):
        id = str(df.loc[i,"ObjectName"])+str(df.loc[i,"Sequence"])+str(df.loc[i,"UserID"])+str(df.loc[i,"Message"])
        unique_id.append(id)
        idx.append(i)
    df["UniqueID"] = unique_id
    df["Index"] = idx
    df.sort_values("TimeTicks",inplace=True,ascending=True)


def filter_dupes_v4(df, time):
    if time<0:
        return df
    uniq_id_index(df)
    print("Filtering duplicate values...")
    keepindex = []
    throwindex = []
    temp_dict = {}
    count = 0
    tot = len(df.loc[:,'UniqueID'])
    for k,t,i in zip(df.loc[:,'UniqueID'],df.loc[:,'TimeTicks'],df.loc[:,'Index']):
        x = (t,i)
        try:
            temp_dict[k].append(x)
        except:
            temp_dict[k] = []
            temp_dict[k].append(x)
    
    bar = Bar("Filtering...",max=tot)
    
    for k in list(temp_dict.keys()):
        ticks_ptr = temp_dict[k][0][0]
        count+=1
        keepindex.append(temp_dict[k][0][1])
        #print(str(count),'/',str(tot),' records processed', end='\r')
        bar.next()
        for i in range(1,len(temp_dict[k])):
            count+=1
            bar.next()
            if (temp_dict[k][i][0] - ticks_ptr)>time:
                keepindex.append(temp_dict[k][i][1])
                ticks_ptr = temp_dict[k][i][0]
            else:
                throwindex.append(temp_dict[k][i][1])
    bar.finish()
    print("")
    print(str(count),'/',str(tot),' records processed')
    print("Records checked. Cleaning up...")
    df.drop(index=throwindex,inplace=True)
    print(len(throwindex)," Duplicate entries removed")

def filter_dupes_v3(df, time):
    if time<0:
        return df
    uniq_id_index(df)
    print("Filtering duplicate values...")
    count = 0
    keepindex = []
    uni_id = []
    start_time = df.loc[0,'TimeTicks']
    start_idx = 0
    end_idx = max(df.index)
    while(True):
        if(start_idx%100 == 0):
            print(str(start_idx),'/',str(end_idx),' records processed', end='\r')
        end_time = start_time+time
        df2 = df[df['TimeTicks']>= start_time]
        df2 = df2[df2['TimeTicks']<end_time]
        uni_id = set(list(df2.loc[:,'UniqueID'].values))
        for u in uni_id:
            keepindex.append(min(df2[df2['UniqueID']==u].index))
        start_idx = max(df[df['TimeTicks']==start_time].index)
        if start_idx + 1 > end_idx:
            break
        start_time = df.loc[start_idx+1,'TimeTicks']
    print("")
    print("Records checked. Cleaning up...")
    keepindex = set(keepindex)
    for i in range(end_idx):
        if i in keepindex:
            continue
        else:
            df.drop(index=i,inplace=True)
            count+=1
    print(count," Duplicate entries removed")
    return df

def filter_dupes_v2(df, time):
    if time<0:
        return df
    uniq_id_index(df)
    print("Filtering duplicate values...")
    count = 0
    max_index = max(df.index)
    df2 = df.copy()
    for i in list(df.loc[:,"Index"].values):
        current_time = df.loc[i,"TimeTicks"]
        current_id = df.loc[i,"UniqueID"]
        if(i%100 == 0):
            print(str(i+1),'/',str(max_index),' records processed', end='\r')
        j=i+1
        if j>max_index:
            break
        next_time = df.loc[j,"TimeTicks"]
        while(next_time <= current_time+time):
            if(current_id == df.loc[j,"UniqueID"]):
                drop_idx = df.loc[j,"Index"]
                try:
                    df2.drop(index=drop_idx,inplace=True)
                    count+=1
                except:
                    pass
            j=j+1
            if j>max_index:
                break
    #df.drop(columns = "UniqueID", inplace = True)
    #df.drop(columns = "Index", inplace = True)
    print(count," Duplicate entries removed")
    return df2

def parse_filter():
    path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"filters.txt"
    try:
        with open(path) as f:
            filt = f.read()
    except:
        print("Error: filters.txt file not found")
        return []
    filter_list = []
    filt=filt.replace('\n',"").replace(" ","").replace("Filter{","").replace("}","").replace("%"," ")
    filt=filt.split(";")
    if len(filt[-1])==0:
        filt.pop()
    for f in filt:
        f = f.split(',')
        f[0] = f[0].replace('FileName=','')
        f[1] = int(f[1].replace('Save=',''))
        f[2] = f[2].replace("(","").replace(")","").replace("[","").replace("]","").split('&')
        f_list = []
        for item in f[2]:
            tup = item.split("=")
            tup[1] = list(tup[1].split("+"))
            tup = tuple(tup)
            f_list.append(tup)
        f_construct = tuple([f[0],f[1],f_list])
        filter_list.append(f_construct)
    return filter_list

def filter_dupes(df):
    #this filters out the duplicate values
    unique_id = []
    for i in range(len(list(df.loc[:,"ObjectName"].values))):
        unique_id = df.loc[i,"ObjectName"]+df.loc[i,"Sequence"]+df.loc[i,"UserID"]
    filt_head = unique_id[0]
    for i in range(1,len(unique_id)):
        if filt_head==unique_id[i]:
            df.drop(index = i, inplace = True)
        else:
            filt_head = unique_id[i]
    

def save_df(df,filename):
    print("Attempting to save file...")
    config_dict = parse_config()
    dest_path = config_dict["Save_Destination_Path"]
    if filename == 'None':
        print('Error: Invalid filename: ',filename)
        sys.exit("Exiting...")
    i=0
    try:
        df.to_csv(dest_path+"\\"+filename+".csv")
    except:

        while(i<20):
            try:
                df.to_csv(dest_path+"\\"+filename+str(i)+".csv")
            except:
                i=i+1
                continue
            break
    if i<20:
        print("Saved successfully")
    else:
        try:
            df.to_csv(filename+".csv")
        except:
            print("Save Failed. Delete old files if any")

def filter_data(df, filters):
    ''' Filter format is like = (FileName = filename.xlsx,Save = 1 or 0,[(Col,[Val1,Val2]),(Col2,[Val1,Val2])])'''
    print("Applying Filter...")
    filename = filters[0]
    save = filters[1]
    for filter in filters[2]:
        col_name = filter[0]
        values_list = filter[1]
        for i in list(df.loc[:,'Index'].values):
            row_ = df.loc[i,col_name]
            if row_ not in values_list:
                df.loc[i,"FilterValue"] = 0
    df = df[df["FilterValue"] == 1]
    if save==1:
        save_df(df,filename)
    return df

def filter_data_v2(df,filters):
    #Efficient but needs to be refined
    print("Applying Filter...")
    filename = filters[0]
    save = filters[1]
    for filter in filters[2]:
        col_name = filter[0]
        val_name = filter[1]
        df = df[df[col_name] == val_name]
    if save==1:
        save_df(df,filename)
    return df

def remove_nan(data_f):
    #Finding the last row of a file. Extra rows with nan entry is removed here. It also combines all the individual data frames 
    #into a single data frame
    #print('Checking for junk values...')
    l = len(data_f)
    junk_frames = []
    for i in range(l):
        if str(data_f.loc[i,'UserAccount']) == 'nan':
            junk_frames.append(i)
    if len(junk_frames)>0:
        # print(len(junk_frames)," junk values found")
        #print("Removing junk frames...")
        data_f.drop(junk_frames,inplace=True)
        #print('Successful.')
    else:
       pass
       # print("No junk frames found.")
       
def parse_config():
    path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"settings.txt"
    try:
        with open(path) as f:
            conf = f.read()
    except:
        print("Error: Settings file not found")
        sys.exit()

    conf = conf.replace(" ","").replace("\n","").replace(';'," ").replace("-{"," ").replace("}","").split()
    conf_dict = {}
    i=0
    while(i<(len(conf))):
        conf_dict[conf[i]] = conf[i+1]
        i=i+2
    return conf_dict

def open_files_of_date(path):
    print("Loading files...")
    f = []
    for(_,_, filenames) in walk(path):
        f.extend(filenames)
        break
    for i in range(6):
        print(f[i])


def update_vocab_file():
    pass


def open_files_in_folder(path, header_no, remove_nan_set=True, read_mes=False):
    #Loading a list of files from a directory
    print("Loading files...")
    f = []
    for (_,_, filenames) in walk(path):
        f.extend(filenames)
        break
    print('Detected ',len(f),' files in the directory')
    dat_frame=[]
    for i in range(len(f)):
        full_path = path+'\\'+f[i]
        #print('Reading File ',f[i])
        try:
            dat_f = pd.read_excel(full_path, index_col = None, header = header_no-1,sheet_name=0, skiprows=0)
            if remove_nan_set:
                remove_nan(dat_f)
            if read_mes:
				hdr = header_no
                while settings_dict["MES_ref_col"] not in list(dat_f.columns):
                    dat_f = pd.read_excel(full_path, index_col = None, header = hdr,sheet_name=0, skiprows=0)
					hrd+=1
        except:
            dat_f = pd.read_csv(full_path)
        dat_frame.append(dat_f)
        print(str(i+1),'/',str(len(f)),' files read successfully', end='\r')
    print("\n")
    print("All files loaded")
    print("Concatinating files...")
    df_concat = pd.concat(dat_frame, ignore_index = True, sort = False)
    print("File concatination sucessful.")  
    return df_concat
#This helps us to open all the excel files in a folder

def sort_by_time(df):
    time_list = []
    shift_list = []
    date = []
    time_value = []
    for i in list(df.loc[:,"EventTime"].values):
        split_val = i.replace(":","-").replace(" ","-").split("-")
        date_time_split = i.split(" ")
        #t_tuple = ((split_val[2]),(split_val[1]),(split_val[0]),(split_val[3]),(split_val[4]),(split_val[5]),0,0,0)
        #print(t_tuple)
        t_tuple = (int(split_val[3]),int(split_val[1]),int(split_val[2]),int(split_val[4]),int(split_val[5]),int(split_val[6]),0,0,0)
        t_ticks = time.mktime(t_tuple)
        time_list.append(t_ticks)
        date.append(date_time_split[1])
        time_value.append(date_time_split[2])
        if (int(split_val[4])>=7) and (int(split_val[4])<15):
            shift_list.append("First Shift")
        elif(int(split_val[4])>=15) and (int(split_val[4])<23):
            shift_list.append("Second Shift")
        else:
            shift_list.append("Night Shift")
    df["Shift"]=shift_list
    df["TimeTicks"]=time_list
    df["Date"] = date
    df["Time"] = time_value

def remove_dupes(df,filter_time_frame):
    pass

def path_extraction(df):
    block_list = [] #Block name that comes after root
    seq_list = [] #individual sequence
    change_type = [] #graphics change or logic related
    equip_group = [] #Equipment group
    end_obj = [] #End object
    application = [] #Application name
    print("Extracting data from path...")
    for i in list(df.loc[:,"Path"].values):
        loc_split = i.replace("]","/").replace("[Location Structure","Graphic Action").replace("[Control Structure","Control Action").replace("SPB_Block","SPB").replace("WBP_Block","WPB").replace("EmulsionBlock","EB").replace("RB_Block","RB").split("/")
        if loc_split[0]=="Control Action":
            try:
                seq_list.append(loc_split[8])
            except:
                seq_list.append('nan')
            
            try:
                block_list.append(loc_split[3])
            except:
                block_list.append('nan')
                
            try:
                change_type.append(loc_split[0])
            except:
                change_type.append(loc_split[0])
                
            try:    
                equip_group.append(loc_split[7])
            except:
                equip_group.append(loc_split[-2])
                
            try:   
                end_obj.append(loc_split[-1])
            except:
                end_obj.append(loc_split[-1])

            try:
                application.append(loc_split[5])
            except:
                application.append('nan')
        else:
            try:
                block_list.append(loc_split[2])
            except:
                block_list.append('nan')
                
            try:
                change_type.append(loc_split[0])
            except:
                change_type.append('nan')
               
            try:
                end_obj.append(loc_split[-1])
            except:
                end_obj.append('nan')
                
            try:
                equip_group.append(loc_split[3])
            except:
                equip_group.append('nan')
            
            application.append('nan')
                
            try:
                seq_list.append(loc_split[-2])
            except:
                seq_list.append('nan')  
                
    print("Data Extraction successful")
    print("Adding to Data Frame")
    df["Block"] = block_list
    df["Change_Type"] = change_type
    df["Application"] = application
    df["Equipment_Group"] = equip_group
    df["Sequence"] = seq_list
    df["Object_Interacted_With"] = end_obj
    print("Successful")
    
def normalize_user_data(df, file_name = 'Employee_details.xlsx'):
    user_id = []
    user_name = []
    user_dept = []
    user_subdept = []
    user_designation = []
    uknown_id = []
    emp_file_path = os.getcwd()+"\\"+"Utility_Files"+"\\"+file_name
    emp_details = pd.read_excel(emp_file_path, index_col = 0, header = 0).transpose()
    emp_dict = emp_details.to_dict(orient = 'series')
    print("Normalizing User Data")
    for i in list(df.loc[:,"UserAccount"].values):
        id_val = i[11:]
        try:
            emp_details_list = list(emp_dict[id_val].values)
        except:
            if id_val not in uknown_id:
                print("Add user details for ",id_val," in the /Utility_Files/Employee_details.xlsx file")
                uknown_id.append(id_val)
            emp_details_list = ['na','na','na','na']
            
        user_id.append(id_val)
        user_name.append(emp_details_list[0])
        user_designation.append(emp_details_list[1])
        user_dept.append(emp_details_list[2])
        user_subdept.append(emp_details_list[3])
        

    df["UserID"] = user_id  
    df["UserName"] = user_name
    df["UserDesignation"] = user_designation
    df["UserDept"] = user_dept
    df["User_Sub_Dept"] = user_subdept  
    print("Successful")
    del uknown_id

def extract_msg(df):
    message = []
    print("Extracting Message Action")
    for i in list(df.loc[:,"Message"].values):
        msg = i.split()
        message.append(msg[0])
    df["Action_Variable"] = message
    print("Successful")

def normalize_object_name(obj_name):
    obj_name = str(obj_name)
    index = re.search("[_]|[-]",obj_name)
    if index is None:
        index = re.search("[0-9]",obj_name)
          
        if index is None:
            split_index = len(obj_name)
        else:
            split_index = index.span(0)[0]
    else:
        split_index = index.span(0)[0]
    if split_index>0:
        return obj_name[0:split_index].lower(), split_index
    else:
        return 'naa', split_index

def obj_vocab_reader(obj_names, file_name = 'Equip_Vocab.xlsx', refresh = True):
    print("Vocab Building started...Reading file...")
    if (refresh==False):
        obj_vocab_dict_temp={}
        obj_equ_list = {"ID":[],"Object":[],"Object Type":[]}
    else:
        obj_vocab_dict_temp={}
        tmp_df = pd.read_excel(os.getcwd()+"\\"+"Utility_Files"+"\\"+file_name,header = 0, index_col = 0)
        id_list = list(tmp_df["ID"].values)
        object_list = list(tmp_df["Object"])
        object_type_list = list(tmp_df["Object Type"])
        obj_equ_list = {"ID":id_list,"Object":object_list,"Object Type":object_type_list}
        for id_val in id_list:
            obj_vocab_dict_temp[id_val] = None
        print("Refresing vocab file...")
    entry_req = False
    for obj_name in obj_names:
        #We try to split the string at _ first. If _ is not found, then we split at 0-9. If that is not found then we
        #retain the entire string
        '''
        index = re.search("[_]|[-]",obj_name)
        if index is None:
            index = re.search("[0-9]",obj_name)
            
            if index is None:
                split_index = len(obj_name)
            else:
                split_index = index.span(0)[0]
        else:
            split_index = index.span(0)[0]'''
        norm_name, _ = normalize_object_name(obj_name)
        '''
        if split_index >0:    
            if obj_name[0:split_index].lower() in obj_vocab_dict_temp:
                pass
            else:
                obj_vocab_dict_temp[obj_name[0:split_index].lower()]=None
                obj_equ_list["ID"].append(obj_name[0:split_index].lower())
                obj_equ_list["Object"].append(obj_name)
                obj_equ_list["Object Type"].append(None)
                entry_req = True
        '''
        if norm_name in obj_vocab_dict_temp:
            pass
        else:
            obj_vocab_dict_temp[norm_name]=None
            obj_equ_list["ID"].append(norm_name)
            obj_equ_list["Object"].append(obj_name)
            obj_equ_list["Object Type"].append(None)
            print("New object entry found : ", norm_name)
            entry_req = True


    if entry_req:
        print("Vocab file needs to be updated...Please make changes in Equip_Vocab.xlsx")
        equip_vocab_df = pd.DataFrame(obj_equ_list)
        equip_vocab_df.to_excel(os.getcwd()+"\\"+"Utility_Files"+"\\"+file_name)
        print("Vocab Building completed...File Saved. Update the vocab file and re-run the program")
        ip = input("Press any key to exit...")
        sys.exit("Exiting")
              
    else:
        print("Vocab file is up to date...")

def object_vocab_file_check(file_name = 'Equip_Vocab.xlsx'):
    
    vocab_df = pd.read_excel(os.getcwd()+"\\"+"Utility_Files"+"\\"+file_name,header=0,index_col=1)
    for i in list(vocab_df.loc[:,"Object Type"]):
        if "nan" == str(i):
            print("Vocab file has incomplete cells. Complete the vocab file before proceeding forward")
            ip = input("Press any key to exit...")
            sys.exit("Exiting...")
    print("Vocab File Checked...Found OK!")
    vocab_df = vocab_df.drop(columns=["Unnamed: 0","Object"]).to_dict()
    return  vocab_df["Object Type"]

def object_type_builder(df, vocab_df, daily_mode = 0):
    object_type = []
    for object_name in list(df.loc[:,"ObjectName"].values):
        norm_obj_name,_ = normalize_object_name(object_name)
        try:
            object_type.append(vocab_df[norm_obj_name])
        except:
            if daily_mode == 1:
                object_type.append('nan')
            else:
                print("Error: Key Error. Vocab file is not up to date.")
                sys.exit("Exiting...")               
    df["Object_Type"] = object_type
    
def action_definition_dict_build_v2(df):
    """
    #Accepts data frame as input
    #Builds a dictionary of dictionaries. Outputs an excel file which allows us to define actions that are acceptable as MI
    """
    action_var = []
    for i in df.loc[:,"Action_Variable"]:
        action_var.append(i)
    
    action_var_set = set(action_var)
    action_var_set = list(action_var_set)
    definition_set = [None for c in range(len(action_var_set))]
    action_dict={"Action_Value":action_var_set,"Def":definition_set}    
    action_df = pd.DataFrame(action_dict)
    action_df.to_excel(os.getcwd()+"\\"+"Utility_Files"+"\\"+"Action_definition_Final.xlsx")
    
def action_definition(df, action_file_path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"Action_definition_Final.xlsx"):
    """
    #This adds action definition to the data frame
    """
    action_def = pd.read_excel(action_file_path, index_col = 1).drop(columns=["Unnamed: 0"])
    action_dict = action_def.to_dict()
    action_class = []
    for item in df.loc[:,"Action_Variable"]:
        try:
            class_type = str(action_dict["Def"][item])
        except KeyError:
            class_type = "NA"
        else:
            if str(class_type)=='nan':
                class_type = "NA"
        
        force_ret = re.findall("Force",item)
        if len(force_ret)>0:
            class_type = "MI"
        action_class.append(class_type)
    return action_class

def read_config():
    logo_path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"config.txt"
    with open(logo_path) as f:
        print(f.read())

def config_reset():
    path = os.getcwd()     
    path = path + "\\"+ "Operator_Action_Files"+"\\"
    df = open_files_in_folder(path,int(settings_dict['Header_Row_Number']))
    extract_msg(df)
    #oper_data_collective.to_excel('All_oper_action_july.xlsx')

    obj_vocab_reader(list(df.loc[:,"ObjectName"].values), refresh = False) #Reset the equip vocab file
    action_definition_dict_build_v2(df) #Only run this if you want to update the action definition
    print("Equipment Vocab file regenerated. Action Definition regenerated.")
    sys.exit("Exiting...")

def admin_mode_check():
    u_name = getpass.getpass(prompt='Username:')
    p_word = getpass.getpass(prompt='Password:')    
    if u_name == "admin" and p_word =="admin":
        config_reset()
    else:
        print("Invalid Credentials...")
        sys.exit("Exiting...")

def resume_temp():
    temp_path = os.getcwd()+"\\"+"temp"+"\\"+"consolidated_dat.pkl"
    temp_path_1 = os.getcwd()+"\\"+"temp"+"\\"+"combined_dat.pkl"
    print("Loading Temp file...")
    try:
        df = pd.read_pickle(temp_path)        
    except:
        try:    
            df = pd.read_pickle(temp_path_1)
        except:
            print("Temp file load error. Check if temp file is available. Else, start new analysis.")
            print("Press any key to exit...")
            _ = input("")
            sys.exit("Exiting...")
        print("Data loaded successfully.")
        return [0,df]
    print("Temp file loaded successfully.")

    try:
        print("Attempting csv file build...")
        df.to_csv("Consolidated_Report.csv")
        print("Consolidated report built successfully. Thanks!")

    except:
        print("Excel file build Failed... Attempting csv file build...")
        try:
            df.to_csv("Consolidated_Report.csv")
            print("Consolidated report built successfully. Thanks!")
        except:
            print("CSV Build also failed... Try running from script")
            sys.exit("Exiting...")
    try:
        os.remove(temp_path)
    except:
        pass
    try:
        os.remove(temp_path_1)
    except:
        pass
    print("Press any key to exit...")
    _ = input("")
    sys.exit("Exiting...")

def print_logo():
    try:
        logo_path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"logo.txt"
        with open(logo_path) as f:
            print(f.read())
    except:
        pass

def detailed_mode():
    print_logo()
    print("Welcome to MI Analysis tool")
    print("Choose your mode of operation")
    print("1. Configuration mode...press 1")
    print("2. User mode...press 2")
    mode = input("")
    if mode == '1':
        admin_mode_check()
    print("_________________________________________________")
    print("User mode selected...Choose execution mode,")
    print("1. Start new analysis")
    print("2. Resume from old analysis")
    print("3. Generate PDF report")
    ret = 1
    mode = input("")
    if mode == '3':
        if settings_dict["report_load_path"]=="current":
            try:
                path = os.getcwd()+"\\"+"Consolidated_Report.csv"
                df = pd.read_csv(path)
            except:
                print("Error reading file. Ensure you have saved the Consolidated_Report.csv in the current folder and try again")
                sys.exit("Exiting...")
        else:
            try:
                df = open_files_in_folder(settings_dict["report_load_path"], header_no=0)
            except:
                print("Error reading file. Ensure you have saved the Reports in the CSV_Report folder and try again. Else use current in settings.")
                sys.exit("Exiting...")
        ret = generateReport_v3(df)
        print("Report generated successfully.")
        sys.exit("Exiting...")

    vocab_path = os.getcwd()+"\\"+"Utility_Files"+"\\"+"Equip_Vocab.xlsx"
    vocab_df = pd.read_excel(vocab_path,header=0,index_col=1).drop(columns=["Unnamed: 0","Object"]).to_dict()
    path = os.getcwd()     
    path = path + "\\"+ "Operator_Action_Files"+"\\"
    temp_path =os.getcwd()+"\\"+"temp"+"\\"+"consolidated_dat.pkl" 
    temp_path_1 =os.getcwd()+"\\"+"temp"+"\\"+"combined_dat.pkl" 
    if mode == '2':
        r = resume_temp()
        ret = r[0]
        df = r[1]
    if ret == 1:
        object_vocab_file_check()
        df = open_files_in_folder(path, int(settings_dict['Header_Row_Number']))
        #oper_data_collective.to_excel('All_oper_action_july.xlsx')
        print("Building temp file checkpoint...")
        df.to_pickle(temp_path_1)
        print("Checkpoint created... If unsuccessful, resume from here.")
        r = resume_temp()
        ret = r[0]
        df = r[1]
    if ret == 0:
        obj_vocab_reader(list(df.loc[:,"ObjectName"].values))
        #action_definition_dict_build_v2(df) #Only run this if you want to update the action definition
        path_extraction(df)
        normalize_user_data(df)
        extract_msg(df)
        sort_by_time(df)
        filter_dupes_v4(df,float(settings_dict['Duplicates_Filter_Time']))
        df.drop(columns = "Message", inplace = True)
        object_vocab_file_check()
        object_type_builder(df,vocab_df["Object Type"])
        action_class = action_definition(df)
        df["Action_Class"] = action_class
        df = df.drop(columns=['Unnamed: 0'])
        if settings_dict["process_allowable_interventions"]=='1':
            df=process_allowable_interventions(df)
        df=df.drop(columns=['index'])
        if settings_dict["Map_MES"]=='1':
            mes_dat_frame = open_files_in_folder(path = os.getcwd()+"\\"+"MES_Reports"+"\\",header_no=int(settings_dict["MES_report_header"]), remove_nan_set = False,  read_mes=True)
            mes_dat_frame = normalize_time_mes(mes_dat_frame)
            mes_map = open_mes_map()
            map_from_mes_data(df,mes_dat_frame,mes_map)
        print("Building temp file checkpoint...")
        df.to_pickle(temp_path)
        print("Checkpoint created... If unsuccessful, resume from here.")
        print("Deleting old temp file...")
        os.remove(temp_path_1)
        print("Delete successful")
        print("Building final report...")
        df.to_csv('Consolidated_Report.csv')
        os.remove(temp_path)
        print("Consolidated report built successfully.")
        if settings_dict['Generate_pdf']=='1':
            df = generateReport_v3(df)
        print("Thanks!\nPress any key to exit")
        _ = input("")
        sys.exit("Exiting...")

def sbt_data_collection():
    vocab_path = settings_dict['Vocab_Path']
    vocab_df = pd.read_excel(vocab_path,header=0,index_col=1).drop(columns=["Unnamed: 0","Object"]).to_dict()
    path = settings_dict['File_Path']
    df = open_files_in_folder(path, int(settings_dict['Header_Row_Number']))
    keep_row = list(np.ones(len(df.loc[:,'ObjectName'].values),dtype=int))
    df["FilterValue"] = keep_row
    path_extraction(df)
    normalize_user_data(df)
    extract_msg(df)
    sort_by_time(df)
    filter_dupes_v4(df,float(settings_dict['Duplicates_Filter_Time']))
    df.drop(columns = "Message",inplace=True)
    object_type_builder(df,vocab_df["Object Type"], daily_mode = 1)
    action_class_path = settings_dict["Action_Class_Path"]
    action_class = action_definition(df,action_file_path=action_class_path)
    df["Action_Class"] = action_class
    df=df.drop(columns=['Unnamed: 0'])
    if settings_dict["process_allowable_interventions"]=='1':
        df=process_allowable_interventions(df)
    if settings_dict["Map_MES"]=='1':
        mes_dat_frame = open_files_in_folder(path = os.getcwd()+"\\"+"MES_Reports"+"\\",header_no=int(settings_dict["MES_report_header"]), remove_nan_set = False,  read_mes=True)
        mes_dat_frame = normalize_time_mes(mes_dat_frame)
        mes_map = open_mes_map()
        map_from_mes_data(df,mes_dat_frame,mes_map)
    filter_list = parse_filter()
    apply_filter(filter_list,df)
    if settings_dict['Generate_pdf']=='1':
        generateReport_v3(df, report_path = settings_dict["Save_Destination_Path"])
       
#Main program starts here
print("Reading settings...")
settings_dict = parse_config()
if settings_dict['Daily_Mode']!='1':
    detailed_mode()
    sys.exit("Exiting...")
else:
    sbt_data_collection()
    
