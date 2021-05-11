#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pandas as pd
import fitz # pip install PyMuPDF
import os


# In[ ]:


def process_apps(filename, appnum_prefix):
    # read CSV file with application info
    df_applications = pd.read_csv(filename, header=0, na_filter=False)
    df_applications.sort_values(by=['End Date'], inplace=True, ascending=True)
    
    appnum_col = []
    num = 1
    
    for i, row in df_applications.iterrows():
        print("---------/nnum: {}, name: {}".format(num, row['Applicant Name']))
        if row['Submit Your Essay Here:'] == '':
            # drop applications without an essay
            df_applications.drop(i, inplace=True)
            print("no essay")
            num += 1
        else:
            ####################
            # rename essay file
            ####################
            appnum_col.append(appnum_prefix + str(num))
            appnum = appnum_col[-1]
            num += 1
            
            #row['Application Number'] = appnum_prefix + str(i+1)
            #appnum = row['Application Number']
            
            old_essay_name = str(row['Respondent ID'])+"_"+row['Submit Your Essay Here:']
            new_essay_name = "{}E.pdf".format(appnum)
            
            try:
                os.rename(os.getcwd()+"/"+old_essay_name, os.getcwd()+"/"+new_essay_name)
            except:
                print("Failed to rename essay {}".format(old_essay_name))
                pass
            
            ##############################
            # merge application and essay
            ##############################
            application = None
            essay = None
            
            try:
                application = fitz.open("{}.pdf".format(appnum))
            except:
                print("Couldn't open application file {}.pdf".format(appnum))
                continue
                
            try:
                essay = fitz.open(new_essay_name)
            except:
                print("Couldn't open essay file {}.pdf".format(new_essay_name))
                continue
                
            try:
                application.insertPDF(essay)
                application.save("{}C.pdf".format(appnum))
                print("Success! Saved {}C.pdf".format(appnum))
                application.close()
                essay.close()
                
            except:
                print("failed to combine app and essay into {}C.pdf".format(appnum))
                if not essay.is_closed:
                    essay.close()
                if not application.is_closed:
                    application.close()
                continue
    # add "Application Number" column
    df_applications['Application Number'] = appnum_col
    # update CSV file
    df_applications.to_csv(filename)
    
### TODO add backup function to save copy of directory before running


# In[ ]:


filenames = ['memorial_applications.csv', 'presidential_applications.csv']

for csv_name in filenames:
    
    # set the applicant number prefix based on scholarship type
    appnum_prefix = 'M'
    if 'presidential' in csv_name:
        appnum_prefix = 'P'
        
    try:
        process_apps(csv_name, appnum_prefix)
        
        # delete all files that don't end in "C.pdf"
        for f_name in os.listdir(os.getcwd()):
            if f_name[0] == appnum_prefix and f_name[-4:]== ".pdf" and f_name[-5] != "C":
                try:
                    os.remove(os.path.join(os.getcwd(),f_name))
                except OSError as e:
                        print("error:" + e.strerror)
    except:
        continue


# In[ ]:




