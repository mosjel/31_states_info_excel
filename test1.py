import pandas as pd
s={"dsd":[1,2,3,4,5],"hamed":["de","de","fgg","fefefw","lllll"],1:[212,3231,5,5565,543]}
s=pd.DataFrame(s)
print(s.iloc[:,1].values)

# print(s[s.iloc[:,2]==3231])
