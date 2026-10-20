import pandas as pd
import numpy as np
import time
from datetime import datetime
def max_comment_prim_date(columns_name):
    comment_date_str=[]
   
    for i in columns_name:
        if ('Комментарии ТУ' in i):
            comment_date_str.append(i[-10:].strip())
       
    
    
    comment_date=pd.to_datetime(comment_date_str,format='%d.%m.%Y')
    
   
    comment_date_cur=comment_date.max() + pd.DateOffset(days=7)
    
    

    return comment_date_cur.strftime('%d.%m.%Y') 

    
def Avans():
    xls_avans=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/МО.xlsx',sheet_name='исх_данные',skiprows=1,usecols=['Номер заказа поставщику','Спецификация (сумма)','Процент авансирования'])
   
    return xls_avans

def LWReport_coments(comment):
    
    if comment.strip()=="":
        return comment
    Date_comment=[]
    Comment_lw=[]
    lines=comment.split('; ')
    for i in lines:
        Date_comment.append(i[:10].strip())
        Comment_lw.append(i[11:].strip())

        

    df=pd.DataFrame({'date_str': Date_comment,'Comment':Comment_lw})
    df['date']=pd.to_datetime(df['date_str'],format='%d.%m.%Y') 
    df_grouped=df.groupby('Comment',sort=False)['date'].agg(['min','max']).reset_index()
    condition=df_grouped['max']!=df_grouped['min']
    df_grouped['Период']=np.where(condition, df_grouped['min'].dt.strftime('%d.%m.%Y') + ' - ' + df_grouped['max'].dt.strftime('%d.%m.%Y'),df_grouped['max'].dt.strftime('%d.%m.%Y'))
    df_clean=df_grouped.dropna(subset=['Период'])
    df_clean_rear=df_clean[['Период','Comment']]
    coments_itog=[]
    for i,j in df_clean_rear.iterrows():
        coments_itog.append(f"{j['Период']}: {j['Comment']}")
    res='\n'.join(coments_itog)
    return  res
def LWReport_melt(prefix='Примечание на'):
    xlsLWReport=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/3LWReport.xlsx',sheet_name='Данные')
    col_1 = ['Заказ поставщику.Номер', 'Планируемая дата отгрузки ТУ',  'Вина поставщика\nДА/НЕТ акт2']
    col_2=[]
    col_skip=["Комментарии_все",'Примечания_все']
    for i in xlsLWReport.columns:
        if i in col_skip:
            continue
        if (prefix in i):
            col_2.append(i)
    report_melted=xlsLWReport.melt(col_1,col_2,"Дата","Текст" )
    
    report_melted['Дата_стр']=report_melted['Дата'].str[-10:].str.strip()
    report_melted['Дата']=pd.to_datetime(report_melted['Дата_стр'], format='%d.%m.%Y', dayfirst=True, errors='coerce')
    report_melted['Текст']=report_melted['Текст'].fillna('')
    report_melted['Текст']=report_melted['Текст'].astype(str)
    report_melted['Планируемая дата отгрузки ТУ']=report_melted['Планируемая дата отгрузки ТУ'].astype(str)
   
    report_melted=report_melted.dropna(subset="Текст")
    report_melted=report_melted[report_melted['Текст']!=""]
    report_melted_grouped=report_melted.groupby(['Заказ поставщику.Номер','Текст'])['Дата'].agg(['min','max','count']).reset_index()
    report_melted_grouped_sorted=report_melted_grouped.sort_values(by="max",ascending=False)
    report_melted_grouped_sorted=report_melted_grouped_sorted.reset_index(drop=True)
    report_melted_grouped_sorted["Период"]=np.where(report_melted_grouped_sorted['min']==report_melted_grouped_sorted["max"],report_melted_grouped_sorted['max'].dt.strftime('%d.%m.%Y')+": "+ report_melted_grouped_sorted['Текст'],report_melted_grouped_sorted['min'].dt.strftime('%d.%m.%Y')+' - '+ report_melted_grouped_sorted['max'].dt.strftime('%d.%m.%Y')+': ' + report_melted_grouped_sorted['Текст'])
    report_melted_grouped_concat=report_melted_grouped_sorted.groupby('Заказ поставщику.Номер')['Период'].agg(";\n".join).reset_index()

    
    
    
    
    return report_melted_grouped_concat
def LWReport():
    xlsLWReport=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/3LWReport.xlsx',sheet_name='Данные')
    col_comments = []
    col_primets = []
    
    col_1 = ['Заказ поставщику.Номер', 'Планируемая дата отгрузки ТУ', '% авансирования', 'Вина поставщика\nДА/НЕТ', 'Наличие в справке', 'Вина поставщика\nДА/НЕТ акт2']
    for i in xlsLWReport.columns:
        if ('Комментарии ТУ' in i) or ('Примечание на' in i):
            col_1.append(i)
        if ('Комментарии ТУ' in i):
            col_comments.append(i)
            
        if ('Примечание на' in i):
            col_primets.append(i)
    
    xlsLWReport = xlsLWReport[col_1]
    #col_1 это список неизменяемых НАЗВАНИЙ столбцов, а также названий столбцов, в которых встречается 'Комментарии ТУ' и 'Примечание на'
    #xlsLWReport['Комментарии_все'] = ''
    
    #for i in col_comments:
     #   xlsLWReport[i]=xlsLWReport[i].fillna("")
      #  xlsLWReport['Комментарии_все'] = np.where(xlsLWReport[i]=="",xlsLWReport['Комментарии_все'],xlsLWReport['Комментарии_все'] + i[-10:] +' '+ xlsLWReport[i] + '; ' )
    #xlsLWReport.to_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Обработка/lwcoments.xlsx')
    

    #xlsLWReport['Комментарии_все']=xlsLWReport['Комментарии_все'].apply(LWReport_coments)
   
    return xlsLWReport 

def UnOtchAvans ():
    xls=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/6UnOtch.xlsx',sheet_name='Лист_1',skiprows=8)
    xls_avans_1=Avans()
    xls['date_str'] = xls['Регистратор'].str.strip().str.split('от').str[1]
    xls['date_avans']=xls['date_str'].str.strip().str.split(' ').str[0]

    xls['date_avans']=pd.to_datetime(xls['date_avans'], dayfirst=True)
    xls['Дата пл.док.']=pd.to_datetime(xls['Дата пл.док.'], dayfirst=True)
    xls['Дата пл.док.']=xls['Дата пл.док.'].fillna(xls['date_avans'])
    #xls['Сумма оплаты']=xls['Сумма оплаты'].astype(int)


    #xls_avans['Спецификация (сумма)']=xls_avans['Спецификация (сумма)'].fillna(0).astype(int)
    xls=xls.sort_values(['Заказ поставщику.Номер','date_avans'])
    xls['rollsum']=xls.groupby('Заказ поставщику.Номер')['Сумма оплаты'].cumsum()

    xls_merged=pd.merge(xls,xls_avans_1,left_on='Заказ поставщику.Номер',right_on='Номер заказа поставщику',how='left')
    xls_merged['Процент авансирования']=xls_merged['Процент авансирования'].fillna(100)

    xls_merged['Процент авансирования дробь']=xls_merged['Процент авансирования']/100
    xls_merged=xls_merged.sort_values(['Заказ поставщику.Номер','date_avans'])
    xls_merged['Оплата от спеки'] = xls_merged['Спецификация (сумма)'] * xls_merged['Процент авансирования дробь']
    xls_merged['оплата аванса'] = np.where(xls_merged['rollsum']>=xls_merged['Оплата от спеки'],'Аванс оплачен','Аванс не оплачен')
    xls_merges_filt=xls_merged[xls_merged['оплата аванса'] == 'Аванс оплачен']
    xls_merges_gr=xls_merges_filt.groupby('Заказ поставщику.Номер')['date_avans'].min().reset_index()
    #xls_merges_gr.to_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Обработка/6UnOtchtestroll.xlsx')  
    #xls_merged.to_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Обработка/6UnOtchtest1.xlsx')  
    return xls_merges_gr
def UO():
    xls=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/1UO.xlsx',sheet_name='Лист_1',skiprows=7)
    xls_drop = xls[~xls['Заказ поставщику'].isna()].reset_index(drop=True)
    xls_drop=xls_drop.iloc[:-1]
    xls_drop['Заказ поставщику']=xls_drop['Заказ поставщику'].str[29:40]
    xls_drop['Тип уведомления']=xls_drop['Уведомление от поставщика'].str[:22].str.strip()
    xls_drop['Дата']=xls_drop['Уведомление от поставщика'].str[64:75].str.strip()
    xls_drop=xls_drop[(xls_drop['Тип уведомления'] =='Извещение об отгрузке') | (xls_drop['Тип уведомления']=='Извещение о готовности')]
    xls_drop['Дата']=pd.to_datetime(xls_drop['Дата'], dayfirst=True)
    xls_drop_gr=xls_drop.groupby(['Заказ поставщику','Тип уведомления'])['Дата'].max().reset_index()
    
    xls_drop_gr_UGO=xls_drop_gr[xls_drop_gr['Тип уведомления'] == 'Извещение о готовности']
    xls_drop_gr_UOO=xls_drop_gr[xls_drop_gr['Тип уведомления'] == 'Извещение об отгрузке']
    #xls_drop_gr.to_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Обработка/xls_drop_gr.xlsx')
    return xls_drop_gr_UGO,xls_drop_gr_UOO

def escuz():
    ESCUZ_4=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/escuzdupl.xlsx',sheet_name='Лист_1')
    FIO_TU=pd.read_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Выгрузки/FIO_TU.xlsx',sheet_name='ФИО-ТУ')
    LWR=LWReport()
    LWR_prim_concat=LWReport_melt()
    LWR_prim_concat=LWR_prim_concat.rename(columns={'Период':'Примечания_все'})
    LWR_comment_concat=LWReport_melt('Комментарии')
    LWR_comment_concat=LWR_comment_concat.rename(columns={'Период':'Комментарии_все'})
    ESCUZ_4['Объект строительства.Группа']=ESCUZ_4.groupby('Заказ поставщику.Номер' )['Объект строительства.Группа'].transform(lambda x: x.ffill().bfill())
    ESCUZ_4_gr=ESCUZ_4.groupby(['Заказ поставщику.Номер','Объект строительства.Группа','Контрагент','№ спецификации','Заказ поставщику.Дата Спецификации','Заказ поставщику.Подразделение организации','Заказ поставщику.Ответственный','Заказ поставщику.Дата поступления','Заказ поставщику.Срок изготовления (в днях)','Заказ поставщику.Оплата','Заказ поставщику.Прибыло','Заказ поставщику.Условие поставки','Заказ поставщику.Объект строительства','Заказ поставщику.Дата отгрузки'],dropna=False)[['Сметная стоимость с НДС','Сумма спецификации с НДС','Оплачено в рыночных ценах (с НДС)','Отгружено поставщиком в ценах поставщика с НДС','Поставлено поставщиком в ценах поставщика с НДС','Количество по спец. ЕНСИ','Количество отгружено в ед. изм. ЕНСИ']].sum().reset_index()
    ESCUZ_GR_TU=pd.merge(ESCUZ_4_gr,FIO_TU,left_on='Заказ поставщику.Ответственный',right_on='Специалист по закупкам',how='left')
    ESCUZ_GR_TU['Заказ поставщику.Срок изготовления (в днях)']=ESCUZ_GR_TU['Заказ поставщику.Срок изготовления (в днях)'].fillna(0)
    conditions_dostavka=[
        ESCUZ_GR_TU['Заказ поставщику.Условие поставки']=='Самовывоз',
        ESCUZ_GR_TU['Заказ поставщику.Объект строительства'].str.contains('Сила Сибири',case=False,na=False),
        ESCUZ_GR_TU['Заказ поставщику.Объект строительства'].str.contains('Харасавэйского',case=False,na=False)]
    choices_dostavka=[0,20,30]
    ESCUZ_GR_TU['Срок Доставки']=np.select(conditions_dostavka,choices_dostavka,default=14)
    mask_izg=(
        ((ESCUZ_GR_TU['Заказ поставщику.Срок изготовления (в днях)']==0)&
        (ESCUZ_GR_TU["Срок Доставки"]==0)) |
        ((ESCUZ_GR_TU['Заказ поставщику.Срок изготовления (в днях)']==0)&
        (ESCUZ_GR_TU['Заказ поставщику.Условие поставки']=='Доставка поставщиком')) 
        
    )
    
    ESCUZ_GR_TU['Заказ поставщику.Срок изготовления (в днях)']=np.where(mask_izg,30,ESCUZ_GR_TU['Заказ поставщику.Срок изготовления (в днях)'])
    ESCUZ_GR_TU_comments=pd.merge(ESCUZ_GR_TU,LWR,left_on='Заказ поставщику.Номер',right_on='Заказ поставщику.Номер',how='left')
    ESCUZ_GR_TU_comments['Процент отгрузки по условиям спецификации']=np.where(ESCUZ_GR_TU_comments['Количество по спец. ЕНСИ']!=0,ESCUZ_GR_TU_comments['Количество отгружено в ед. изм. ЕНСИ']/ESCUZ_GR_TU_comments['Количество по спец. ЕНСИ'],0)
    ESCUZ_GR_TU_comments['Процент отгрузки по есцуз']=np.where(ESCUZ_GR_TU_comments['Сумма спецификации с НДС']!=0,ESCUZ_GR_TU_comments['Отгружено поставщиком в ценах поставщика с НДС']/ESCUZ_GR_TU_comments['Сумма спецификации с НДС'],0)
    ESCUZ_GR_TU_comments_prim=pd.merge(ESCUZ_GR_TU_comments,LWR_prim_concat,left_on='Заказ поставщику.Номер',right_on='Заказ поставщику.Номер',how='left')
    ESCUZ_GR_TU_comments_prim_comm=pd.merge(ESCUZ_GR_TU_comments_prim,LWR_comment_concat,left_on='Заказ поставщику.Номер',right_on='Заказ поставщику.Номер',how='left')
    #ESCUZ_4_gr.to_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Обработка/ESCUZ_4tes.xlsx',index=False)  
    return ESCUZ_GR_TU_comments_prim_comm
def Escuz_date_avanse():
    escuz_orig=escuz()
    LastDateAvans=UnOtchAvans()
    MO_avanse=Avans()
    UGO,UOO = UO()
    today=pd.Timestamp.today().normalize()
    
    
    MO_avanse['Процент авансирования']=MO_avanse['Процент авансирования'].fillna(100)

    LastDatePayAvans=pd.merge(escuz_orig,LastDateAvans,left_on='Заказ поставщику.Номер',right_on='Заказ поставщику.Номер',how='left')
    LastDatePayAvansUgo=pd.merge(LastDatePayAvans,UGO,left_on='Заказ поставщику.Номер',right_on='Заказ поставщику',how='left')
    LastDatePayAvansUgoUo=pd.merge(LastDatePayAvansUgo,UOO,left_on='Заказ поставщику.Номер',right_on='Заказ поставщику',how='left')
    LastDatePayAvansUgoUoAvans=pd.merge(LastDatePayAvansUgoUo,MO_avanse,left_on='Заказ поставщику.Номер',right_on='Номер заказа поставщику',how='left')
    
    LastDatePayAvansUgoUoAvans['Обьект для отчета']=LastDatePayAvansUgoUoAvans['Объект строительства.Группа']+"."+LastDatePayAvansUgoUoAvans['Заказ поставщику.Объект строительства']
    LastDatePayAvansUgoUoAvans['Расчетная дата отгрузки']=LastDatePayAvansUgoUoAvans['date_avans']+pd.to_timedelta(LastDatePayAvansUgoUoAvans['Заказ поставщику.Срок изготовления (в днях)'],unit='D')+pd.to_timedelta(LastDatePayAvansUgoUoAvans['Срок Доставки'],unit='D')
    LastDatePayAvansUgoUoAvans['Оплачено в рыночных ценах (с НДС)']=LastDatePayAvansUgoUoAvans['Оплачено в рыночных ценах (с НДС)'].fillna(0)
    LastDatePayAvansUgoUoAvans['Отгружено поставщиком в ценах поставщика с НДС']=LastDatePayAvansUgoUoAvans['Отгружено поставщиком в ценах поставщика с НДС'].fillna(0)
    LastDatePayAvansUgoUoAvans['Неисполненные обязательства']=np.where(LastDatePayAvansUgoUoAvans['Оплачено в рыночных ценах (с НДС)']-LastDatePayAvansUgoUoAvans['Отгружено поставщиком в ценах поставщика с НДС']<0,0,LastDatePayAvansUgoUoAvans['Оплачено в рыночных ценах (с НДС)']-LastDatePayAvansUgoUoAvans['Отгружено поставщиком в ценах поставщика с НДС'])
    LastDatePayAvansUgoUoAvans['Процент авансирования']=LastDatePayAvansUgoUoAvans['Процент авансирования']/100
    LastDatePayAvansUgoUoAvans['Наличие оплаты']=np.where(LastDatePayAvansUgoUoAvans['Оплачено в рыночных ценах (с НДС)']>=(LastDatePayAvansUgoUoAvans['Сумма спецификации с НДС']*LastDatePayAvansUgoUoAvans['Процент авансирования']),'Оплачено','Не оплачено')
    LastDatePayAvansUgoUoAvans['Наличие неисполненных обязательств']=np.where(LastDatePayAvansUgoUoAvans['Неисполненные обязательства']<100,'Неисполненные обязательства отвсутсвуют','Наличие неисполненных обязательств')
    LastDatePayAvansUgoUoAvans['Факт просрочки отгрузки']=np.where(LastDatePayAvansUgoUoAvans['Расчетная дата отгрузки']>today,'Не просрочено','Просрочено')
    conditions_prim=[
        LastDatePayAvansUgoUoAvans['Сумма спецификации с НДС']<1000000,
        LastDatePayAvansUgoUoAvans['Заказ поставщику.Прибыло']=='Прибыло',
        (LastDatePayAvansUgoUoAvans['Заказ поставщику.Оплата']=='Не оплачен') & (LastDatePayAvansUgoUoAvans['Процент отгрузки по условиям спецификации']<1),
        (LastDatePayAvansUgoUoAvans['Процент отгрузки по условиям спецификации']==1) & (LastDatePayAvansUgoUoAvans['Процент отгрузки по есцуз']==1),
        LastDatePayAvansUgoUoAvans['Процент отгрузки по условиям спецификации']>1,
        ((LastDatePayAvansUgoUoAvans['Расчетная дата отгрузки']> (today + pd.DateOffset(months=2))) & (LastDatePayAvansUgoUoAvans['Расчетная дата отгрузки'].notna())),
        ((today<=LastDatePayAvansUgoUoAvans['date_avans']+pd.DateOffset(days=20)) & (LastDatePayAvansUgoUoAvans['date_avans'].notna()))

    ]
    choices_prim=['Менее миллиона по спецификации','Поставлено по данным есцуз','Не оплачено','Отгружено по всем показателям','Отгружено более 100 процентов','Расчетный срок отгрузки не наступил','Крайняя дата платежа ок']
    date_cur_comment=max_comment_prim_date(LastDatePayAvansUgoUoAvans.columns)
    
    LastDatePayAvansUgoUoAvans[f"Примечание на {date_cur_comment}"]=np.select(conditions_prim,choices_prim,"")
    #LastDatePayAvansUgoUoAvans["Примечание"]=np.select(conditions_prim,choices_prim,"")
    LastDatePayAvansUgoUoAvans[f'Комментарии ТУ Актуальные на {date_cur_comment}']=''
    
    columns_sample=[
    "Заказ поставщику.Номер",
    "Объект строительства.Группа",
    "Заказ поставщику.Объект строительства",
    "Обьект для отчета",
    "Контрагент",
    "№ спецификации",
    "Заказ поставщику.Дата Спецификации",
    "Заказ поставщику.Подразделение организации",
    "Заказ поставщику.Ответственный",
    "Наименоване товарного управления",
    "Заказ поставщику.Дата поступления",
    "Заказ поставщику.Срок изготовления (в днях)",
    "Заказ поставщику.Оплата",
    "Заказ поставщику.Прибыло",
    "Сметная стоимость с НДС",
    "Сумма спецификации с НДС",
    "Оплачено в рыночных ценах (с НДС)",
    "Отгружено поставщиком в ценах поставщика с НДС",
    "Поставлено поставщиком в ценах поставщика с НДС",
    "Количество по спец. ЕНСИ",
    "Количество отгружено в ед. изм. ЕНСИ",
    "Заказ поставщику.Условие поставки",
    "Дата_y",
    "Срок Доставки",
    "date_avans",
    "Процент отгрузки по условиям спецификации",
    "Процент отгрузки по есцуз",
    "Дата_x",
    "Расчетная дата отгрузки",
    "Процент авансирования",
    "Неисполненные обязательства",
    "Наличие оплаты",
    "Наличие неисполненных обязательств",
    "Факт просрочки отгрузки",
    "Планируемая дата отгрузки ТУ",
    "% авансирования",
    "Вина поставщика\nДА/НЕТ",
    "Наличие в справке",
    "Вина поставщика\nДА/НЕТ акт2",
    f'Комментарии ТУ Актуальные на {date_cur_comment}',
    "Комментарии ТУ Актуальные на 12.12.2025",
    "Комментарии ТУ Актуальные на 05.12.2025",
    "Комментарии ТУ Актуальные на 28.11.2025",
    "Комментарии_все",
    "Примечания_все",
    
    "Примечание на на 28.11.2025",
    "Примечание на на 05.12.2025",
    "Примечание на 12.12.2025",
    f"Примечание на {date_cur_comment}"
]

    return LastDatePayAvansUgoUoAvans[columns_sample]


if __name__ == '__main__':
 start=time.perf_counter()
 df=Escuz_date_avanse()

 df.to_excel(r'C:/Users/Samski/Documents/Мониторинг поставщиков/Обработка/melt_test.xlsx')
 
 end=time.perf_counter()
 print(f"Время работы: {end - start:.4f} сек")
