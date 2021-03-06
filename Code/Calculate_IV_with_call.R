if(T==0) return(c(0,0,0))
ftr<-function(x){
  d1=(log(S/K)+(r+0.5*x^2)*T)/(x*sqrt(T))
  d2=d1-x*sqrt(T)
  return(S*pnorm(d1)-K*exp((-r*T))*pnorm(d2)-C)
}

tar=0
for(i in c(-1,0,1,2,3,4,5,6,7,8,9)){
  for(j in 1:9){
    if(ftr(tar+j*10^(-i))==0){
      tar=tar+j*10^(-i)
      break
    }
    else if(ftr(tar+(j+1)*10^(-i))==0){
      tar=tar+(j+1)*10^(-i)
      break
    }
    else if(ftr(tar+j*10^(-i))*ftr(tar+(j+1)*10^(-i))<0){
      tar=tar+j*10^(-i)
    }
    else if(j==9){
      if(ftr(tar+j*10^(-i))*ftr(tar+(j+1)*10^(-i))>0){
        tar=tar+0
      }
    }
  }
}
nd1=(log(S/K)+(r+0.5*tar^2)*T)/(tar*T^(1/2))
return(c((tar*100),pnorm(nd1),((dnorm(nd1))/(S*tar*(T^(1/2))))))