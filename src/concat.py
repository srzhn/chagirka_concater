import pandas as pd

def join_cells(df: pd.DataFrame, sep: str='_'):
    """
    Соединяет значения ячеек экселя. 
    """
    labels = []
    for row in df.values:
        labels.append(sep.join(
            list(map(str,
                     filter(lambda x: not pd.isna(x), row)))))
        
    labels = pd.Series(labels, name='label')
    return labels

def join_cells_with_colnames(df: pd.DataFrame, sep: str='_'):
    """
    Соединяет значения ячеек экселя c названиями колонок (если имеются).
    """
    labels = []
    cols = df.columns
    unnamed_cols = list(map(lambda x: x.startswith('Unnamed'), cols))
    
    for row in df.values:
        # Заводим список для формирования будущей строки
        s = []
        for i in range(len(row)):
            # пропускаем, если ячейка пустая
            if pd.isna(row[i]) or len(str(row[i]))==0:
                continue

            # если есть название колонки - добавляем в строку
            if not unnamed_cols[i]:
                s.append(cols[i])
                
            # добавляем значение в строком формате
            s.append(str(row[i]))
        
        # формируем строку и добавляем в итог
        labels.append(sep.join(s))
    
    
    labels = pd.Series(labels, name='label')
    return labels
