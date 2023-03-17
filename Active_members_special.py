import pandas as pd

import numpy as np

import re as regex
[2]
users = pd.read_excel('../Active Members/Active Members Updated March 2023.xlsx')

users.head()

[3]
users.info()
<class 'pandas.core.frame.DataFrame'>
RangeIndex: 35658 entries, 0 to 35657
Data columns (total 7 columns):
 #   Column                                           Non-Null Count  Dtype         
---  ------                                           --------------  -----         
 0   User Id                                          35658 non-null  int64         
 1   Username                                         35658 non-null  object        
 2   Profile Url                                      35658 non-null  object        
 3   Email                                            35656 non-null  object        
 4   Last Activity Date (UTC)                         35658 non-null  datetime64[ns]
 5   Last Activity Date (Coordinated Universal Time)  35658 non-null  datetime64[ns]
 6   Total Activity Count                             35658 non-null  int64         
dtypes: datetime64[ns](2), int64(2), object(3)
memory usage: 1.9+ MB

[4]
users.describe()

[5]
users.drop(['Profile Url', 'Last Activity Date (Coordinated Universal Time)'], inplace=True, axis=1)

users.head()

[6]
users.columns = ['User ID', 'Username', 'Email', 'Date', 'Total Activity Count']

users.head()

[7]
users['Date'] =pd.to_datetime(users['Date'])
[8]
users['month_year'] = users['Date'].apply(lambda x: x.strftime('%B-%Y'))
[9]
users.head()

[10]
internal = users[users['Email'].str.contains('.*@sdl', '.*@rws', na=False)]

internal

[11]
internal = internal.astype({'month_year': 'datetime64'})

internal.info()
<class 'pandas.core.frame.DataFrame'>
Int64Index: 4666 entries, 6625 to 35656
Data columns (total 6 columns):
 #   Column                Non-Null Count  Dtype         
---  ------                --------------  -----         
 0   User ID               4666 non-null   int64         
 1   Username              4666 non-null   object        
 2   Email                 4666 non-null   object        
 3   Date                  4666 non-null   datetime64[ns]
 4   Total Activity Count  4666 non-null   int64         
 5   month_year            4666 non-null   datetime64[ns]
dtypes: datetime64[ns](2), int64(2), object(2)
memory usage: 255.2+ KB

[12]
internal.describe()

[13]
members = internal.groupby(['month_year']).agg(
    monthly_users=('User ID', 'count')
)

members

[17]
writer = pd.ExcelWriter('Active Members Special.xlsx',
                       engine='xlsxwriter',
                       engine_kwargs={'options': {'strings_to_urls':False}}
                       )
members.to_excel(writer)

writer.close()