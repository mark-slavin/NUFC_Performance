# NUFC_Performance
https://sites.google.com/view/mark-slavin/projects/excel

Excel based analysis on the performance of Newcastle United against the previous 5 seasons.

It all started with a query to Google's Gemini LLM
I told it which columns I wanted and after a to and frow, when it only game me eight games per season so I had to ask it for all of the games (it applogised).

I was presented with the results for each season in seperate tables.
<img width="790" height="662" alt="image" src="https://github.com/user-attachments/assets/04e94ccf-3b6f-447a-bcb1-e43e3f9c6160" />

I added a column to convert the W/L/D into points using the Switch function (=SWITCH([@Result],"w",3,"d",1,"l",0) ).
<img width="527" height="185" alt="image" src="https://github.com/user-attachments/assets/79fe0b5d-a68e-4b74-90c7-18c2d0891bef" />

I then added points total column to create a running total for each season.
<img width="690" height="137" alt="image" src="https://github.com/user-attachments/assets/833a556c-1b96-41b5-b824-f59732d5e593" />

Using these new columns, I was able to produce this chart. The more recent seasons are darker greys; the current season is red for contrast.
<img width="649" height="324" alt="image" src="https://github.com/user-attachments/assets/573a0571-88f8-4987-9530-f9073d9d8a1c" />

Using the formula 

=UNIQUE(DataTable[[Opponent]:[H/A]]) 

to create a list of teams and fixtures and

 =TOROW(UNIQUE(DataTable[Season])) 

to create a list of season column headers, I made this table.
<img width="751" height="209" alt="image" src="https://github.com/user-attachments/assets/1c9e6dc8-1b6f-40bb-a444-272ef695791f" />

The formula to return the points for each fixture is

=IF(COUNTIFS(DataTable[Season],'Vs Team'!C$1,DataTable[Opponent],'Vs Team'!$A2,DataTable[H/A],'Vs Team'!$B2)>0,SUMIFS(DataTable[Points],DataTable[Season],'Vs Team'!C$1,DataTable[Opponent],'Vs Team'!$A2,DataTable[H/A],'Vs Team'!$B2),"") 

This seems overly complicated, but it only returns a score if a game took place. This will account for teams that weren't in the PL that season or for games in the current season that haven't been played yet.

="Over that last 6 seasons, "&COUNTA(UNIQUE(A:A))-1&" different teams have played in the Premier league" 

<img width="208" height="117" alt="image" src="https://github.com/user-attachments/assets/4a8afff7-2283-4f8a-91cb-7e8fb304046a" />

="Against teams that they have played this season (so far) and last season, Newcastle United are "&ABS(SUM(J:J))&" points "&IF(SUM(J:J)<0,"worse","better")&" off than last season" 

<img width="190" height="136" alt="image" src="https://github.com/user-attachments/assets/74d7b7af-2fc6-4401-a05f-faea47a07c8a" />

https://sites.google.com/view/mark-slavin/projects/excel





