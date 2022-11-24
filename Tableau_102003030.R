#installation of readxl package to read excel files
install.packages("readxl")
library("readxl")
library(dplyr)

#For Table 1
table1<- read_excel("C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\Table1.xlsx")
g<-data.frame(table1)  #storing data into g as data frame
View(g)

#Replacing NA values with Mean
g$Geometry.Coordinates.1.0[is.na(g$Geometry.Coordinates.1.0)]<-mean(g$Geometry.Coordinates.1.0,na.rm=TRUE)

g_flight<- group_by(g,Flight)

df_mean<-summarise(g_flight,passengers=mean(No_of_Passengers, na.rm = TRUE))

for(x in df_mean$Flight)
{
  a<-is.na(g$No_of_Passengers)
  b<-g$Flight==x 
  g$No_of_Passengers[a&b]=floor(df_mean$passengers[df_mean$Flight==x])
}
View(g)

#For Properties Table
table2<- read_excel("C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\Table2.xlsx")
p<-data.frame(table2)  #storing data into p as data frame
View(p) #Before preprocessing
#Preprocessing of properties table-->

#Replacing NA values with Mean
t<-mean(p$Properties.Flysfo.Base.Flight.Number,na.rm=TRUE)
p$Properties.Flysfo.Base.Flight.Number[is.na(p$Properties.Flysfo.Base.Flight.Number)]<-trunc(t)

#Declaring function to calculate Mode
Mode<-function(x){
  ux<-na.omit(unique(x))
  tab<-tabulate(match(x,ux)); ux[tab==max(tab)]
}

#Replacing NA values with Mode
p$Properties.Edtf.Inception[is.na(p$Properties.Edtf.Inception)]<-Mode(p$Properties.Edtf.Inception)

p$Properties.Flysfo.Airline[is.na(p$Properties.Flysfo.Airline)]<-Mode(p$Properties.Flysfo.Airline)

p$Properties.Flysfo.Base.Airline[is.na(p$Properties.Flysfo.Base.Airline)]<-Mode(p$Properties.Flysfo.Base.Airline)

p$Properties.Flysfo.Event[is.na(p$Properties.Flysfo.Event)]<-Mode(p$Properties.Flysfo.Event)

p$Properties.Flysfo.Flight.Number[is.na(p$Properties.Flysfo.Flight.Number)]<-Mode(p$Properties.Flysfo.Flight.Number)

p$Properties.Flysfo.Gate[is.na(p$Properties.Flysfo.Gate)]<-Mode(p$Properties.Flysfo.Gate)

View(p) #After preprocessing


#For Route Table
table3<- read_excel("C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\Table3.xlsx")
r<-data.frame(table3)  #storing data into r as data frame
View(r)

#Spliting the Route column into two as "From" and "To" using stringr
library(stringr)
s<-str_split_fixed(r$Route, "-", 2)
colnames(s)=c("From","To")
d<-cbind.data.frame(r,s)
View(d)

#Merging of Tables
f<-merge(g,p,by="SrNo")
air_travel<-merge(f,d,by="SrNo")
View(air_travel)

#Formatting Date Column
air_travel$Date<-format(as.POSIXct(air_travel$Date), format="%d-%m-%y")
View(air_travel)

#Putting Dataframe in Excel File
install.packages("writexl")
library("writexl")
write_xlsx(air_travel,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\air_travel.xlsx")

#Query 1
x<-g[,c(1,8)]
write_xlsx(x,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q1.xlsx")

#Query 2
y<-g[,c(1,8)]
z<-d[,c(1,2)]
f<-merge(y,z,by="SrNo")
f$Date<-format(as.POSIXct(f$Date), format="%d-%m-%y")
write_xlsx(f,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q2.xlsx")

#Query 3
y<-g[,c(1,8)]
z<-d[,c(1,2)]
f<-merge(y,z,by="SrNo")
f$Date<-format(as.POSIXct(f$Date), format="%d-%m-%y")
write_xlsx(f,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q3.xlsx")

#Query 4
y<-g[,c(1,8)]
busy<-d[,c(1,3,4,5)]
f<-merge(y,busy,by="SrNo")
write_xlsx(f,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q4.xlsx")

#Query 5
distance<-g[,c(1:5)]
busy<-d[,c(1,3,4,5)]
f<-merge(distance,busy,by="SrNo")
write_xlsx(f,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q5.xlsx")

#Query 6
distance<-g[,c(1:5)]
busy<-d[,c(1,3,4,5)]
x<-g[,8]
names(x)=c("Flight")
b<-cbind(busy,x)
f<-merge(distance,b,by="SrNo")
write_xlsx(f,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q6.xlsx")

#Query 7
passengers<-g[,c(1,7,8)]
route<-d[,c(1,3)]
f<-merge(passengers,route,by="SrNo")
write_xlsx(f,"C:\\Users\\MRITUNJAY\\Desktop\\Tableau Project\\q7.xlsx")