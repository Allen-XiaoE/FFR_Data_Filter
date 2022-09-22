from datetime import timedelta,date
import os
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

def dataprocess(robottype,path):
    os.chdir(path)
    #1.Setting File path
    robotfilepath = robottype + '.xlsm'
    warrantyfilepath = 'Warranty Parts Information List.xlsx'
    failurecategorypath = 'Failure Category.xlsx'
    dataresourcefilepath = '{0}_DataResource.xlsx'.format(robottype)
    isexistfile = os.path.exists(dataresourcefilepath)

    #2.Get File
    robotfile = pd.read_excel(robotfilepath,sheet_name = None,keep_default_na=False)
    WarrantyData = pd.read_excel(warrantyfilepath,sheet_name='query (3)',keep_default_na=False)#DataFrame
    FailureCategoryData = pd.read_excel(failurecategorypath,sheet_name='Sheet1',keep_default_na=False)#DataFrame
    if isexistfile:
        dataresourcefile = pd.read_excel(dataresourcefilepath,sheet_name=None ,keep_default_na=False)
        Old_FailureData = dataresourcefile['Failure Data']
        Old_DMAICData = dataresourcefile['DMAIC Data']
        Old_DMAICData['D-Date'] = pd.to_datetime(Old_DMAICData['D-Date'])
        Old_DMAICData['M-Date'] = pd.to_datetime(Old_DMAICData['M-Date'])
        Old_DMAICData['A-Date'] = pd.to_datetime(Old_DMAICData['A-Date'])
        Old_DMAICData['I-Date'] = pd.to_datetime(Old_DMAICData['I-Date'])
        Old_DMAICData['C-Date'] = pd.to_datetime(Old_DMAICData['C-Date'])

    #3.Get DataFrame and DataFrame Filter
    DeliveryData = robotfile[robottype]
    DeliveryData = DeliveryData.iloc[[2,5,6,12,13,14,18],7:].T
    DeliveryData = DeliveryData.reset_index()
    DeliveryData.columns = ['Delivery Month','Delivery Volume','Volume accumulated','Failures accumulated','FFR 18 month rolling',
                            'Volume 18 month rolling','Failures 18 month rolling','Failures delivery month']
    DeliveryData.insert(1,'Delivery Date',pd.to_datetime(DeliveryData['Delivery Month'],format='%Y%m'))

    QMData = robotfile['GRP QM Data']
    #QMData.set_index('QM No.',inplace=True)
    New_FailureData = QMData[['QM No.','QM Month','QM Date','Delivery Month','Delivery date','Group','Item','Defect Type','Failure description',
    'PRU ','Analysis description','TestCenter Destination','Customer','Duty Counter','Art. No','Robot S/N','Cause description','Cause category']]
    New_FailureData.columns = ['QM No.','QM Month','QM Date','Delivery Month','Delivery Date','Group','Item','Defect Type','Failure Description',
    'PRU','Analysis description','Testcenter Destination','Customer','Duty Counter','Art No.','Robot S/N','Cause description','Cause category']
    New_FailureData[['Position','Failure Item','Failure Type','Axis','Failure','Status','Root Cause Type']] = ''
    New_FailureData['Status'] = 'CL'
    New_FailureData['QM No.'] = New_FailureData['QM No.'].astype('str')
    New_FailureData['Analysis description'] = New_FailureData['Analysis description'].astype('str')
    New_FailureData['Cause category'] = New_FailureData['Cause category'].astype('str')
    New_FailureData = New_FailureData.replace(['-',','],'0')
    New_FailureData.loc[New_FailureData['Cause category'] == '0','Status'] = 'WM'
    New_FailureData.sort_values(by=['QM Date','QM No.'],inplace=True,ascending=[False,False])
    FailureCategoryData = FailureCategoryData[['Position','Item','Failure Type','Axis']]


    #4.填充数据
    for qmno in New_FailureData['QM No.']:
        #填充旧数据
        if isexistfile:
            Old_FailureData['QM No.'] = Old_FailureData['QM No.'].astype('str')
            #Get Data From Old FailureData Sheet
            olddatarow = Old_FailureData[Old_FailureData['QM No.'] == qmno]
            if not olddatarow.empty:
                New_FailureData.loc[New_FailureData['QM No.']==qmno,'Root Cause Type'] = olddatarow['Root Cause Type'].values[0]
                New_FailureData.loc[New_FailureData['QM No.']==qmno,'Position'] = olddatarow['Position'].values[0]
                New_FailureData.loc[New_FailureData['QM No.']==qmno,'Failure Item'] = olddatarow['Failure Item'].values[0]
                New_FailureData.loc[New_FailureData['QM No.']==qmno,'Failure Type'] = olddatarow['Failure Type'].values[0]
                New_FailureData.loc[New_FailureData['QM No.']==qmno,'Axis'] = olddatarow['Axis'].values[0]
                New_FailureData.loc[New_FailureData['QM No.']==qmno,'Failure'] = olddatarow['Failure'].values[0]
        
        #Get Data From warranty Sheet
        warrantydatarow = WarrantyData[WarrantyData['QM No'] == qmno]
        if not warrantydatarow.empty:
            if warrantydatarow['QM Status'].values[0] == 'Closed':
                New_FailureData.loc[New_FailureData['QM No.'] == qmno,'Status'] = 'CL'
                New_FailureData.loc[New_FailureData['QM No.'] == qmno,'Root Cause Type'] = warrantydatarow['Failure Cause Summary'].values[0]
        
            #Get Analysis Description from warrantyexcel.detail
            if New_FailureData.loc[New_FailureData['QM No.'] == qmno,'Analysis description'].values[0] == '0' and warrantydatarow['Detail'].values[0] != '':
                New_FailureData.loc[New_FailureData['QM No.'] == qmno,'Analysis description'] = warrantydatarow['Detail'].values[0]

    #5.统计Failure的数量
    past_18_months = date.today() - timedelta(days=18*30)
    past_18_months = pd.to_datetime(past_18_months)
    New_DMAICData_past_18_months = New_FailureData[(New_FailureData['QM Date'] > past_18_months)&(New_FailureData['Delivery Date'] > past_18_months)]['Failure']
    New_DMAICData = New_FailureData.drop_duplicates(['Failure'])
    New_DMAICData = New_DMAICData[New_DMAICData['Failure'] != '']['Failure']
    New_DMAICData = pd.DataFrame(New_DMAICData)
    New_DMAICData['Quantity'] = 0
    New_DMAICData[['DMAIC','D-Date','M-Date','A-Date','I-Date','C-Date','Record']] = ''

    New_DMAICData.columns = ['Failure','Quantity','DMAIC','D-Date','M-Date','A-Date','I-Date','C-Date','Record']
    #New_DMAICData = New_DMAICData.reset_index()

    if not New_DMAICData_past_18_months.empty:
        #New_DMAICData = pd.DataFrame(New_DMAICData.value_counts())
        # New_DMAICData = New_DMAICData.reset_index()
        # New_DMAICData[['DMAIC','D-Date','M-Date','A-Date','I-Date','C-Date','Record']] = ''
        # New_DMAICData.columns = ['Failure','Quantity','DMAIC','D-Date','M-Date','A-Date','I-Date','C-Date','Record']
        for i in New_DMAICData_past_18_months.value_counts().keys():
            New_DMAICData.loc[New_DMAICData['Failure'] == i,'Quantity'] = New_DMAICData_past_18_months.value_counts()[i]
        
        #填充旧数据
    if isexistfile:
        for failure in New_DMAICData['Failure']:
            old_failurevalue = Old_DMAICData[Old_DMAICData['Failure'] == failure]
            if not old_failurevalue.empty:
                columnlist = ['DMAIC','D-Date','M-Date','A-Date','I-Date','C-Date','Record']
                for cols in columnlist:
                    New_DMAICData.loc[New_DMAICData['Failure'] == failure,cols] = old_failurevalue[cols].values[0]
    
    New_DMAICData.sort_values(by=['Quantity','Failure'],inplace=True,ascending=[False,True])
    New_DMAICData.index = pd.RangeIndex(start=1,stop =len(New_DMAICData)+1 ,step=1)
    New_DMAICData['D-Date'] = pd.to_datetime(New_DMAICData['D-Date'])
    New_DMAICData['M-Date'] = pd.to_datetime(New_DMAICData['M-Date'])
    New_DMAICData['A-Date'] = pd.to_datetime(New_DMAICData['A-Date'])
    New_DMAICData['I-Date'] = pd.to_datetime(New_DMAICData['I-Date'])
    New_DMAICData['C-Date'] = pd.to_datetime(New_DMAICData['C-Date'])
    #6.Failure Catogery 排序
    # for col in FailureCategoryData.columns:
    #     FailureCategoryData[col] = FailureCategoryData[col].sort_values(ascending=True,inplace=False)
    #6.Save Data Into Excel
    newpath = '{0}_DataResource.xlsx'.format(robottype)
    writer = pd.ExcelWriter(newpath,engine='xlsxwriter')
    New_FailureData.to_excel(writer,sheet_name='Failure Data',index = False)
    DeliveryData.to_excel(writer,sheet_name='Delivery Data',index=False)
    New_DMAICData.to_excel(writer,sheet_name='DMAIC Data',index=True)
    FailureCategoryData.to_excel(writer,sheet_name='Failure Category',index=False)
    writer.save()
    print('{0} DataProcess is Done!'.format(robottype))
    # print('---------------------------------\n')


