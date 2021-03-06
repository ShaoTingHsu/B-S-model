if(g==1){
  x1<-read.xlsx2(file="RR.xlsx",sheetIndex=1,startRow = 9,header = F,encoding = "UTF-8")
  x1date<<-as.Date(sub("日期：",replacement="",as.character(read.xlsx(file="RR.xlsx",sheetIndex=1,startRow = 1,header = F,encoding = "UTF-8")[2,1])))
  x1=x1[-which(x1$X2==""),]
  # x1=x1[-which(is.na(x1$X2)),]
  #將不同到期日、call、put分開
  sum1=0
  sum11=0
  sum2=0
  sum22=0
  sum3=0
  sum33=0
  sum4=0
  sum44=0
  sum5=0
  sum55=0
  for(i in 1:(nrow(x1))){
    if(as.character(x1[i,2])==as.character(x3[1,2])){
      if(as.character(x1[i,4])=="Call"){
        sum1=sum1+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum11=sum11+1
      }
      
    }
    if(as.character(x1[i,2])==as.character(x3[2,2])){
      if(as.character(x1[i,4])=="Call"){
        sum2=sum2+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum22=sum22+1
      }
    }
    if(as.character(x1[i,2])==as.character(x3[3,2])){
      if(as.character(x1[i,4])=="Call"){
        sum3=sum3+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum33=sum33+1
      }
    }
    if(as.character(x1[i,2])==as.character(x3[4,2])){
      if(as.character(x1[i,4])=="Call"){
        sum4=sum4+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum44=sum44+1
      }
    }
    if(as.character(x1[i,2])==as.character(x3[5,2])){
      if(as.character(x1[i,4])=="Call"){
        sum5=sum5+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum55=sum55+1
      }
    }
    
  }
  temp=cons.matrix(sum1,sum11,1,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],1,x1date,x1)
  temp=cons.matrix(sum2,sum22,2,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],2,x1date,x1)
  temp=cons.matrix(sum3,sum33,3,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],3,x1date,x1)
  temp=cons.matrix(sum4,sum44,4,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],4,x1date,x1)
  temp=cons.matrix(sum5,sum55,5,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],5,x1date,x1)
  
}
else if(g==2){
  x1<-read.xlsx2(file="RR.xlsx",sheetIndex=2,startRow = 9,header = F,encoding = "UTF-8")
  x1[,9]<-x1[,8]
  x1[,15]<-x1[,12]
  x1date<<-as.Date(sub("日期：",replacement="",as.character(read.xlsx(file="RR.xlsx",sheetIndex=2,startRow = 1,header = F,encoding = "UTF-8")[2,1])))
  x1=x1[-which(x1$X2==""),]
  x1<-x1[-which(x1$X9=="-"),]
  #將不同到期日、call、put分開
  sum6=0
  sum66=0
  sum7=0
  sum77=0
  sum8=0
  sum88=0
  for(i in 1:(nrow(x1))){
    if(as.character(x1[i,2])==as.character(x3[6,2])){
      if(as.character(x1[i,4])=="Call"){
        sum6=sum6+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum66=sum66+1
      }
      
    }
    if(as.character(x1[i,2])==as.character(x3[7,2])){
      if(as.character(x1[i,4])=="Call"){
        sum7=sum7+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum77=sum77+1
      }
    }
    if(as.character(x1[i,2])==as.character(x3[8,2])){
      if(as.character(x1[i,4])=="Call"){
        sum8=sum8+1
      }
      if(as.character(x1[i,4])=="Put"){
        sum88=sum88+1
      }
    }
  }
  temp=cons.matrix(sum6,sum66,6,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],6,x1date,x1)
  temp=cons.matrix(sum7,sum77,7,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],7,x1date,x1)
  temp=cons.matrix(sum8,sum88,8,x1)
  getoutcome(temp[,1:6],temp[,-c(1:6)],8,x1date,x1)
}