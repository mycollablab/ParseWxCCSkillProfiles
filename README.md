# ParseWxCCSkillProfiles

Script will parse export of Skill Profiles from WxCC and provide a matrix of profiles and skill and then take that same matrix and create a CSV file for upload to WxCC.

Only requires openpyxl for excel. 

```
pip3 install openpyxl
```

Takes argument of the filename. If CSV it assumes you want to create an Excel matrix, if XLSX it assumes you want a CSV import file.
```
python3 parseSkillProfiles.py skillsExport.csv
...
python3 parseSkillProfiles.py skillsExport.xlsx
```
