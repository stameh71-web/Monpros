import pandas as pd

import time
import os
import sqlite3
from MonPos import TRANFORM_DIR,TU_DIR,Sql_Base_dir,max_comment_prim_date





def zapros_ty():
    
    Zapros_TY = pd.read_excel(TRANFORM_DIR/'отчет для анализа.xlsx', sheet_name='Sheet1')
  
    Date_max=(pd.to_datetime(max_comment_prim_date(Zapros_TY.columns),dayfirst=True)).strftime('%d.%m.%Y')
    Zapros_TY_filt=Zapros_TY[Zapros_TY[f'Примечание на {Date_max}'].isin(['Запросить ТУ','Ручная проверка'])]

    columns_sample = [
        "Дата отчета",
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
        "Срок Доставки",
        "date_avans",
        "Процент отгрузки по условиям спецификации",
        "Процент отгрузки по есцуз",
        "Расчетная дата отгрузки",
        "Процент авансирования",
        "Планируемая дата отгрузки ТУ",
        "% авансирования",
        "Вина поставщика ДА/НЕТ акт2",
        f'Комментарии ТУ Актуальные на {Date_max}',
        'Комментарии_все']
    
    Zapros_TY_filt=Zapros_TY_filt[columns_sample]
    Zapros_TY_filt.to_excel(TU_DIR/'Запросы ТУ.xlsx', index=False)    
    TU_name=['УОМТР','УТМ','УТО','УЭКИП','УСП','УВП']
    result={}
    for tu in TU_name:
        chunk=Zapros_TY_filt[Zapros_TY_filt['Наименоване товарного управления']==tu]
        chunk=chunk[columns_sample]
        chunk["Планируемая дата отгрузки ТУ"]=chunk["Планируемая дата отгрузки ТУ"].dt.strftime('%d.%m.%Y')
        chunk["Вина поставщика ДА/НЕТ акт2"]=chunk["Вина поставщика ДА/НЕТ акт2"].replace(['НЕТ','нет','Нет'],'')
        chunk.to_excel(TU_DIR/f'Запросы ТУ {tu}.xlsx', index=False)
        result[tu]=chunk
    return Zapros_TY_filt[columns_sample], TU_name , Date_max

if __name__ == '__main__':
    start = time.perf_counter() 
    For_TU ,TU,max_date= zapros_ty()
    conn=sqlite3.connect(Sql_Base_dir)
    cur=conn.cursor()
    cur.execute("DELETE FROM ZaprosTU where \"Дата отчета\" =?",(max_date,))
    For_TU.to_sql(
        name='ZaprosTU',
        con=conn,
        if_exists='replace',
        index=False

    )
    deleted=cur.rowcount
    print(f'ZaprosTU: удалено {deleted}, добавлено {len(For_TU)}')
    conn.commit()
    conn.close()   

    end = time.perf_counter()
    print(f"Время работы: {end - start:.4f} сек")

    for tu in TU:
        os.startfile(TU_DIR/f'Запросы ТУ {tu}.xlsx')