# msa
Material from my presentation at Make Data Smart 2019 - Trondheim: https://event.dnd.no/mds2019td/

In this repo you will find a cheat sheet for making reports in MS Word and MS PowerPoint using the Python libraries: python-docx and python-pptx

## Data -- keep your data close and know what you have
* Survey 
* Data from your Data Ware House
* Eksternal data sources:
(SSB, NAV, NVE, Yr, Nordpool, AtB, Ruter, Statens vegvesen) 

## Analytics -- four kinds of analytical tools:
* Descriptive statistics - focus on visualization
* Statistical models: Causal  and statistical relations
* Empirical analyis: Compare apples with apples – Models such as «Propensity score matching», «Synthetic control» and «Differece-in-Difference» are good when one are lacking natural experiments
* Machine learning: Suited when one are making predictions or when there are to many explanatory variables

## Presentation of the results -- which tools is appropiate to use for presentation?
1.  Cut and paste results into Word or PowerPoint
2.	Python to Word or PowerPoint -> python-docx and Python-pptx 
3.	Dashboard tools: (PowerBI/Tablau/VA)
4.	Jupyter Notebook: jupyter nbconvert --to html notebook.ipynb
5.  Interactive plots like bqplot together with Voila! 
