# PowerAuditor
![](https://raw.githubusercontent.com/1mm0rt41PC/PowerAuditor/master/bin/logo.png) Powered by  [![Excel o365](https://upload.wikimedia.org/wikipedia/commons/thumb/e/ed/Microsoft_Office_Excel_%282013%E2%80%93present%29.svg/101px-Microsoft_Office_Excel_%282013%E2%80%93present%29.svg.png)](https://www.office.com)

# Presentation
PowerAuditor is an excel sheet over-vitaminated with macro. PowerAuditor allows you :
- to write Pentest reports in a minimum of time.
- to no longer write the same information twice
- share your vulnerability sheets with your co-worker via a git server
- to share your knowledge with your team via a git server

# First install for the infrastructure
Require:
- Windows 10
- Excel / Word upper to 2013
- a git repository for the reports template
- a git repository for the vulnerabilities template


```batch
C:\Users\1mm0rt41PC> git clone https://github.com/1mm0rt41PC/PowerAuditor.git
C:\Users\1mm0rt41PC> cd PowerAuditor

C:\Users\1mm0rt41PC\PowerAuditor> cd template
C:\Users\1mm0rt41PC\PowerAuditor\template> git init .
C:\Users\1mm0rt41PC\PowerAuditor\template> :: Put here your template (xlsm and docx) with a name like xxxx_v1-EN.xlsm and xxxx_v1-EN.docx (See Example_v1-FR.xlsm and Example_v1-FR.docx)
C:\Users\1mm0rt41PC\PowerAuditor\template> git add .
C:\Users\1mm0rt41PC\PowerAuditor\template> git commit -am "Init"
C:\Users\1mm0rt41PC\PowerAuditor\template> git remote add origin xxx@xxxx.fr:xxxx/myRepo-for-template.git
C:\Users\1mm0rt41PC\PowerAuditor\template> git push -u origin master

C:\Users\1mm0rt41PC\PowerAuditor\template> cd ..\vulndb
C:\Users\1mm0rt41PC\PowerAuditor\vulndb>:: In this folder will be store all vulnerability that will be shared
C:\Users\1mm0rt41PC\PowerAuditor\template> git init .
C:\Users\1mm0rt41PC\PowerAuditor\template> git add .
C:\Users\1mm0rt41PC\PowerAuditor\template> git commit -am "Init"
C:\Users\1mm0rt41PC\PowerAuditor\template> git remote add origin yyyy@yyyy.fr:xxxx/myRepo-for-vuln.git
C:\Users\1mm0rt41PC\PowerAuditor\template> git push -u origin master
```

# Usage
1) Copy the `PowerAuditor_last.xlsm` in your pentest project.
2) Create a folder `vuln` forlder and create a subfolder for each vulnerability

```
.
├── PowerAuditor_last.xlsm
└── vuln
    ├── Citrix vulnerable to IKEExt
    │   ├── proof 1.png
    │   ├── This is a HTTP request.http
    │   └── proof 2.png
    ├── Clear text communication
    │   └── proof.png
    └── XSS
        └── proof.png
```

3) Run `PowerAuditor_last.xlsm` and enable macro
4) Select a `Report type` and a Language
5) Go in the new sheet (ie: Example_v1-EN)
6) In the ribbon tab `PowerAuditor`, click on `Fill excel with proof` to fill the the excel with all vuln from the folder `vuln`.
7) Fill all lines about your vulnerabilities
8) To export all theses datas to the word document click on `Export Excel to Word`


### Development

Use the file `dev\PowerAuditor.xlsm`
