# v0114
import pandas as pd
from tkinter import *

root=Tk()
root.title("Plotter - v0114")

xax=Label(root, text="X Axis Column number:")
e_xax=Entry(root, width=30)

yax1=Label(root, text="Y Axis Column number:")
e_yax1=Entry(root, width=30)

fltr=Label(root, text="Filter Column number:")
e_fltr=Entry(root, width=30)

colmns=Label(root, text="Number of Charts Columns:")
e_colmns=Entry(root, width=30)

current_state=Label(root, text="-")
note=Label(root, text="Input file must be named '4.xlsx'.",fg='red')
notex=Label(root, text="To plot multiple KPIs, split column number with ;",fg='red')
status = Label(root, text = "Version: 0114 - Coded by: Ahmad Dawara", bd=2, relief=SUNKEN, anchor = E)

# ----------------------------------------------------

s_xax='a'
s_yax1='a'
s_fltr='a'
s_colmns='a'
def gene():
    s_xax=e_xax.get()
    s_yax1=e_yax1.get()
    s_fltr=e_fltr.get()
    s_colmns=e_colmns.get()

    df = pd.read_excel('4.xlsx', 'Sheet1')

    if len(s_yax1)>1:
        y_list=s_yax1.split(';')
        print(y_list)

        n_list=[]

        for cn in y_list:
            cn=int(cn)
            cn=cn-1
            n_list.append(df.columns[int(cn)])

        print(n_list) 

        xx=int(s_xax)-1
        ff=int(s_fltr)-1

        x_col_name=df.columns[xx]
        f_col_name=df.columns[ff]

        df.sort_values(by=[str(f_col_name)],inplace=True)

        cells_list=df[str(f_col_name)].unique()

        excel_file = 'cell-kpi-all-1.xlsx'
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        
        #df_append=pd.DataFrame(columns = [str(s_xax),str(s_yax1),str(s_yax2)])
        appended_data = list()
        dfs_sizes=list()
        
        sz=0
        counter=1
        reset_count=1

        for cel in cells_list:
            
            dft=df[df[str(f_col_name)]==cel]
            df1 = dft[[str(x_col_name)]]
            df2 = dft[dft.columns.intersection(n_list)]
            dff = dft[str(f_col_name)]
            
            sz=sz+len(df1)
            dfs_sizes.append(sz)
            df_temp=pd.concat([dff,df1,df2],axis=1)
            df_temp.sort_values(by=[str(x_col_name)],inplace=True)
            #print(df_temp.head())
            appended_data.append(df_temp)
        
        app_df = pd.concat(appended_data,axis=0)
        reset_count=reset_count+1
        print(dfs_sizes)
        #app_df.to_csv('app_df.csv',index=False)
        app_df.to_excel(writer, sheet_name='Data', index=False)
        app_df.to_csv('x.csv')
        #print(counter)

        if counter==1:
            workbook= writer.book
            worksheet = workbook.add_worksheet('charts')
        #-------------------------------------------------------------------
        i=0
        r=2
        x=1
        y=1
        yl=0
        size_ind=0
        r=2
        for n in n_list:
            #yl=0
            size_ind=0
            
            for cel in cells_list:
                
                #size_ind=0
                #worksheet = writer.sheets['charts']
                chart = workbook.add_chart({'type': 'line'})
                # categories --- x-axies
                # Configure the series of the chart from the dataframe data.
                print('********')
                print(size_ind)
                if(size_ind==0):
                    chart.add_series({'name': str(n),'categories': ['Data', 1, 1, dfs_sizes[size_ind], 1], 'values': ['Data', 1, r, dfs_sizes[size_ind], r]})
                    size_ind=size_ind+1
                    print('first')
                
                elif(size_ind==len(dfs_sizes)-1):
                    chart.add_series({'name': str(n),'categories': ['Data', dfs_sizes[size_ind-1]+1, 1, len(app_df), 1], 'values': ['Data', dfs_sizes[size_ind-1]+1, r, len(app_df), r]})
                    
                    #size_ind=size_ind+1
                    print('last')
                        
                else:
                    chart.add_series({'name': str(n),'categories': ['Data', dfs_sizes[size_ind-1]+1, 1, dfs_sizes[size_ind], 1], 'values': ['Data', dfs_sizes[size_ind-1]+1,  r, dfs_sizes[size_ind],  r]})
                    size_ind=size_ind+1
                    print('mid')

                chart.set_title({'name':str(cel)})
                chart.set_x_axis({'text_axis': True})
                chart.set_y_axis({
                    'major_gridlines': {
                        'visible': True,
                        'line': {'color': '#D9D9D9'}
                    },
                })
                chart.set_legend({'position': 'bottom'})
                s_colmns=int(s_colmns)
                if(i%s_colmns!=0):
                    x=x+8
                else:
                    x=1

                if(i%s_colmns==0 and i!=0):
                    y=y+15

                # Insert the chart into the worksheet
                worksheet.insert_chart(y,x, chart)
                #print(i,j)
                print('----')
                
                i=i+1
            r=r+1
            yl=yl+1
        counter=counter+1       
            # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        current_state.config(text='Done.')
    

    else:    
        xx=int(s_xax)-1
        yy=int(s_yax1)-1
        ff=int(s_fltr)-1

        x_col_name=df.columns[xx]
        y_col_name=df.columns[yy]
        f_col_name=df.columns[ff]
        
        df.sort_values(by=[str(x_col_name)],inplace=True)
        cells_list=df[str(f_col_name)].unique()

        excel_file = 'cell-kpi-all-1.xlsx'
        writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
        
        #df_append=pd.DataFrame(columns = [str(s_xax),str(s_yax1),str(s_yax2)])
        appended_data = list()
        dfs_sizes=list()

        sz=0
        for cel in cells_list:
            #df_temp=pd.DataFrame(columns = [str(s_xax)])
            dft=df[df[str(f_col_name)]==cel]
            df1 = dft[[str(x_col_name)]]
            df2 = dft[str(y_col_name)]
            dff = dft[str(f_col_name)]

            sz=sz+len(df1)
            dfs_sizes.append(sz)
            df_temp=pd.concat([df1,df2,dff],axis=1)
            #print(df_temp.head())
            appended_data.append(df_temp)

        app_df = pd.concat(appended_data,axis=0)

        print(dfs_sizes)
        #app_df.to_csv('app_df.csv',index=False)
        app_df.to_excel(writer, sheet_name='Data', index=False)

        workbook= writer.book
        worksheet = workbook.add_worksheet('charts')
    #-------------------------------------------------------------------
        i=0
        j=1
        x=1
        y=1
        size_ind=0

        for cel in cells_list:
            #worksheet = writer.sheets['charts']
            chart = workbook.add_chart({'type': 'line'})
            # categories --- x-axies
            # Configure the series of the chart from the dataframe data.
            if(size_ind==0):
                chart.add_series({'name': str(y_col_name),'categories': ['Data', 1, 0, dfs_sizes[size_ind], 0], 'values': ['Data', 1, j, dfs_sizes[size_ind], j]})
            
            elif(size_ind==len(dfs_sizes)-1):
                chart.add_series({'name': str(y_col_name),'categories': ['Data', dfs_sizes[size_ind-1]+1, 0, len(app_df), 0], 'values': ['Data', dfs_sizes[size_ind-1]+1, j, len(app_df), j]})
                print(len(app_df))
                    
            else:
                
                chart.add_series({'name': str(y_col_name),'categories': ['Data', dfs_sizes[size_ind-1]+1, 0, dfs_sizes[size_ind], 0], 'values': ['Data', dfs_sizes[size_ind-1]+1, j, dfs_sizes[size_ind], j]})
            
                print(dfs_sizes[size_ind-1]+1)

            chart.set_title({'name':str(cel)})
            chart.set_x_axis({'text_axis': True})
            chart.set_y_axis({
                'major_gridlines': {
                    'visible': True,
                    'line': {'color': '#D9D9D9'}
                },
            })
            chart.set_legend({'position': 'bottom'})
            s_colmns=int(s_colmns)
            if(i%s_colmns!=0):
                x=x+8
            else:
                x=1

            if(i%s_colmns==0 and i!=0):
                y=y+15

            # Insert the chart into the worksheet
            worksheet.insert_chart(y,x, chart)
            print(i,j)
            print('----')
            size_ind=size_ind+1
            i=i+1
            #j=j+1
                
        # Close the Pandas Excel writer and output the Excel file.
        writer.save()

        current_state.config(text='Done.')

# ----------------------------------------------------

b=Button(root, text="Generate File", command=gene)
# --------------------------------------------------------------------

xax.grid(row = 0, column = 0, pady=10, padx=5)
e_xax.grid(row = 0, column = 1, pady=10, padx=5)

yax1.grid(row = 1, column = 0, pady=10, padx=5)
e_yax1.grid(row = 1, column = 1, pady=10, padx=5)

fltr.grid(row = 3, column = 0, pady=10, padx=5)
e_fltr.grid(row = 3, column = 1, pady=10, padx=5)

colmns.grid(row = 4, column = 0, pady=10, padx=5)
e_colmns.grid(row = 4, column = 1, pady=10, padx=5)

current_state.grid(row = 5, column = 1, pady=10, padx=5)

b.grid(row = 6, column = 1, pady=5)  
note.grid(row = 7, column = 1,  padx=5)
notex.grid(row = 8, column = 1,  padx=5)
status.grid(row=10, column=0, columnspan=4, sticky=W+E)

# -------------------------------------------------------------------

root.mainloop()


