O.C.W<-matrix(numeric(3*nrow(C.)),ncol=3)
O.P.W<-matrix(numeric(3*nrow(P.)),ncol=3)



for(i in 1:nrow(C.)){
  O.C.W[i,]=cdgv(C.[i,6],C.[i,1],C.[i,4],C.[i,5],C.[i,2])
}
for(i in 1:nrow(P.)){
  O.P.W[i,]=pdgv(P.[i,6],P.[i,1],P.[i,4],P.[i,5],P.[i,2])
}


Delta.W<-numeric(nrow(O.C.W)) #Delta(加權)相似值
diffDelta.W<-numeric(nrow(O.C.W)) #Delta(加權)差異
relativeK.W<-numeric(nrow(O.C.W)) #對應履約價(加權)
relativeWC.W<-numeric(nrow(O.C.W)) #對應未沖銷契約量(加權)
Gamma.W<-numeric(nrow(O.C.W)) #相似值對應的Gamma(加權)
IV.W<-numeric(nrow(O.C.W)) #相似值對應的IV(加權)
diffIV.W<-numeric(nrow(O.C.W)) #IV差(加權)
turn.diffIV.W<-numeric(nrow(O.C.W)) #逆轉(現貨價)
minindex.W=seq(1,by=0,length=nrow(O.C.W))
for(i in 1:nrow(O.C.W)){
  ocpw=abs(O.P.W[1,2]+O.C.W[i,2])
  for(j in 2:nrow(O.P.W)){
    if(ocpw>abs(O.P.W[j,2]+O.C.W[i,2])){
      ocpw=abs(O.P.W[j,2]+O.C.W[i,2])
      minindex.W[i]=j
    }
  }
  relativeK.W[i]=P.[(minindex.W[i]),1]
  relativeWC.W[i]=P.[(minindex.W[i]),3]
  Delta.W[i]=O.P.W[minindex.W[i],2]
  Gamma.W[i]=O.P.W[minindex.W[i],3]
  IV.W[i]=O.P.W[minindex.W[i],1]
  diffDelta.W[i]=O.C.W[i,2]-Delta.W[i]
  diffIV.W[i]=O.C.W[i,1]-IV.W[i]
}

for(i in 1:(nrow(O.C.W)-1)){
  if(diffIV.W[i]*diffIV.W[(i+1)]<0){
    turn.diffIV.W[i]=1
    turn.diffIV.W[(i+1)]=1
  }
}

maxind=max(nrow(O.C.W),nrow(O.P.W))
final<-array(numeric(9*maxind),c(maxind,9))
if(name<=5){
  colnames(final) <- c("Delta相似值","Delta差異","對應履約價","對應未沖銷契約量","相似值對應的Gamma","相似值對應的IV","IV差","逆轉(現貨價)",as.character(x3[name,6]))
}
if(name>5){
  colnames(final) <- c("Delta相似值","Delta差異","對應履約價","對應成交量","相似值對應的Gamma","相似值對應的IV","IV差","逆轉(現貨價)",as.character(x3[name,6]))
}
for(i in 1:nrow(O.C.W)){
  final[i,1]<-Delta.W[i] #Delta(加權)相似值
  final[i,2]<-diffDelta.W[i] #Delta(加權)差異
  final[i,3]<-relativeK.W[i] #對應履約價(加權)
  final[i,4]<-relativeWC.W[i] #對應履約價(加權)
  final[i,5]<-Gamma.W[i] #相似值對應的Gamma(加權)
  final[i,6]<-IV.W[i] #相似值對應的IV(加權)
  final[i,7]<-diffIV.W[i] #IV差(加權)
  final[i,8]<-turn.diffIV.W[i] #逆轉(現貨價)
  final[i,9]<-C.[i,6] #加權指數每日收盤價
}

call.data<-array(numeric(7*maxind),c(maxind,7))
if(name<=5){
  colnames(call.data) <- c("日期","到期月份(週別)","CALL履約價","CALL未沖銷契約量","CALL.Delta","CALL.Gamma","CALL.IV(%)")
}
if(name>5){
  colnames(call.data) <- c("日期","到期月份(週別)","CALL履約價","CALL成交量","CALL.Delta","CALL.Gamma","CALL.IV(%)")
}
call.data[,1]=as.character(x1date)#日期
call.data[,2]=as.character(x3[name,3]) #到期月份(週別)
for(i in 1:nrow(O.C.W)){
  call.data[i,3]=C.[i,1] #CALL履約價
  call.data[i,4]=C.[i,3] #CALL未沖銷契約量
  call.data[i,5]=O.C.W[i,2] #CALL.Delta
  call.data[i,6]=O.C.W[i,3] #CALL.Gamma
  call.data[i,7]=O.C.W[i,1] #CALL.IV(%)
  
}
put.data<-array(numeric(5*maxind),c(maxind,5))
if(name<=5){
  colnames(put.data) <- c("PUT履約價","PUT未沖銷契約量","PUT.Delta","PUT.Gamma","PUT.IV(%)")
}
if(name>5){
  colnames(put.data) <- c("PUT履約價","PUT成交量","PUT.Delta","PUT.Gamma","PUT.IV(%)")
}
for(i in 1:nrow(O.P.W)){
  put.data[i,1]=P.[i,1] #PUT履約價
  put.data[i,2]=P.[i,3] #"PUT未沖銷契約量
  put.data[i,3]=O.P.W[i,2] #PUT.Delta
  put.data[i,4]=O.P.W[i,3] #PUT.Gamma
  put.data[i,5]=O.P.W[i,1] #PUT.IV(%)
}

put.data=put.data[order(put.data[,1],decreasing = T),]
termenal<-cbind(call.data,put.data,final)
if(name==1){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==2){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==3){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==4){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==5){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==6){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==7){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}
if(name==8){
  write.xlsx(termenal,file=paste0(x1date,"報表.xlsx"),sheetName = paste0(as.character(x3[name,1]),"_",as.character(x3[name,7])),col.names = T, row.names = F, append = TRUE, showNA = TRUE)
}