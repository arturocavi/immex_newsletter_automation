#Obtención de Datos para Gráficas y Texto de IMMEX para Boletín
#Fuente: INEGI, IMMEX Datos Abiertos Segmento Manufacturero y No Manufacturero por entidad federativa

#Limpiar variables
rm(list=ls())


########## LIBRERIAS ##########
library(dplyr)
library(openxlsx)
library(lubridate)


########## DIRECTORIO ##########
#MODIFICACIÓN DE DATOS AUTOMÁTICA (la que se usa siempre, a menos que el mes en que se corra no coincida con el mes de publicación de INEGI)
#Determinar el mes y el año de la encuesta, bajo el supuesto que se publica dos meses después
#Hacer adecuaciones en caso de que se ejecute el código en una fecha distinta a la publicación de INEGI (ver "MODIFICACIÓN DE DATOS MANUAL" 4 líneas abajo)
mes = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-2 months")[2],6,7))
año = as.numeric(substr(seq(Sys.Date(), length = 2, by = "-2 months")[2],1,4))
#A mano
#mes=mes-1

#Ponerle cero al mes2
if (mes < 10) {
  mes2 = paste0("0",mes)
  # Si el mes es menor a 10, le pegamos un 0 delante.
} else {
  mes2 = as.character(mes)
}

directorio = paste0("C:/Users/arturo.carrillo/Documents/IMMEX/",año," ",mes2)
setwd(directorio)


########## PERIODOS ##########
#Fecha para bases de csv
fcsv=paste(año,mes2,sep = "_")

#Periodos de titulares de graficas
meses = data.frame(c("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre"))
colnames(meses) = "MES" #Cambia nombre de la columna

periodo1=paste0("enero 2013-",paste(meses[mes,1],año))
periodo2=paste(meses[mes,1],año)

#Contador para saber el numero de datos que necesitas dependiendo del mes
#Ej: 2019: ENE=12, FEB=13, MAR=14, ABR=15, MAY=16, JUN=17, JUL=18, AGO=19, SEP=20, OCT=21, NOV=22, DIC=23,
#    2020: ENE=24, FEB=25, MAR=26, ABR=27, MAY=28, JUN=29, JUL=30, AGO=31, SEP=32, OCT=33, NOV=34, DIC=34, ...
c=(año-2018)*12+mes-1


########## NÚMEROS CÁRDINALES ##########
#Números cardinales DEL 1 AL 20 en masculino
num_cardinales = c("primer","segundo","tercer","cuarto","quinto","sexto","séptimo","octavo","noveno","décimo","decimoprimer","decimosegundo","decimotercer","decimocuarto","decimoquinto","decimosexto","decimoséptimo","decimoctavo","decimonoveno","vigésimo")


########## DESCARGAR BASES DE INEGI ##########
#Sector Manufacturero
temp = tempfile() #crear archivo temporal
#Descargar archivo del INEGI de Manufacturas:
download.file("https://www.inegi.org.mx/contenidos/programas/immex/datosabiertos/immex_manu_ent_csv.zip",temp)

for(y in 2009:año){
  nombre_base_csv=paste0("immex_manu_ent_",y,".csv")
  
  if(y==2009){
    base_csv_manu = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv)),encoding = "UTF-8")
  }else{
    base_csv_manus = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv)),encoding = "UTF-8")#año siguiente
    base_csv_manu=rbind(base_csv_manu,base_csv_manus)
  }
}
rm(base_csv_manus)
unlink(temp) #borrar archivo  temporal (el descargado de INEGI)


#Sector No Manufacturero
temp = tempfile() #crear archivo temporal
#Descargar archivo del INEGI de No Manufacturas:
download.file("https://www.inegi.org.mx/contenidos/programas/immex/datosabiertos/immex_nomanu_ent_csv.zip",temp)

for(y in 2009:año){
  nombre_base_csv=paste0("immex_nomanu_ent_",y,".csv")
  
  if(y==2009){
    base_csv_nomanu = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv)),encoding = "UTF-8")
  }else{
    base_csv_nomanus = read.csv(unz(temp,paste0("conjunto_de_datos/",nombre_base_csv)),encoding = "UTF-8")#año siguiente
    base_csv_nomanu=rbind(base_csv_nomanu,base_csv_nomanus)
  }
}
rm(base_csv_nomanus)
unlink(temp) #borrar archivo  temporal (el descargado de INEGI)


########## SEGMENTO MANUFACTURERO ##########
#Leer base
#Cifras en miles de pesos
#nombre_base_csv=paste0("immex_manu_ent_2009_01_",fcsv,".csv")
#base_csv=read.csv(nombre_base_csv)
base_csv=base_csv_manu

#Crear variable de mes con ceros
nrb=nrow(base_csv)
m0=rep(NA,nrb)
for (i in 1:nrb){
  if (base_csv$ID_MES[i] < 10){
    m0[i]=paste0("0",base_csv$ID_MES[i])
  }
  else{
    m0[i]=base_csv$ID_MES[i]
  }
}

#Crear variable de fecha con anio y mes
#FECHA=paste(base_csv$X.U.FEFF.ANIO,m0,sep="-")
FECHA=paste(base_csv$ANIO,m0,sep="-")


#Base de IMMEX No Manufacturero con fecha (variables seleccionadas)
varsec=cbind.data.frame(FECHA,base_csv)
varsec=select(varsec,-ANIO,-ID_MES,ENT=ID_ENTIDAD,-ID_MUNICIPIO,EST=NUM_DE_ESTAB_ACT,-G210A,-M310B,-M710B,-M999B,-starts_with("K"),-ESTATUS)

#Variables de totales
#Total Trabajadores (Obreros y Administrativos; Contratados y Subcontratados)
TRA=varsec$H114A+varsec$H200A+varsec$I400A+varsec$I500A
#Total Salarios (Obreros y Administrativos; Contratados)
SAL=varsec$J114A+varsec$J200A
#Total Horas Trabajadas (Obreros y Administrativos; Contratados y Subcontratados)
HOR=varsec$H114D+varsec$H200D+varsec$I400D+varsec$I500D
#Total Exportaciones (Productos, Maquila y Otros)
EXP=varsec$M310C+varsec$M710C+varsec$M999C

#Variables de Seguridad Social y Prestaciones (variables seleccionadas ademas de totales)
#Seguridad Social
SEG=varsec$J114A
#Prestaciones
PRE=varsec$J200A

#Base con totales y variables seleccionadas
base_tot=cbind.data.frame(varsec[1:3],TRA,SAL,SEG,PRE,HOR,EXP)


#Lista de entidades en el programa immex (manufactureros)
e=unique(varsec$ENT)

#Colapsar datos de cada entidad por fecha
for (i in e){
  #Datos de Entidad
  datos_ent=base_tot[base_tot$ENT==i,]
  datos_ent=select(datos_ent,-ENT)
  #Colapsar por mes todos los municipios del estado
  colaps=aggregate(.~FECHA, datos_ent, sum)
  colaps=mutate(colaps,ENT=i)
  colaps=colaps[,c(9,1:8)]
  if (i==e[1]){
    datos=colaps
  }else{
    datos=rbind.data.frame(datos,colaps)
  }
}

#Identificador Entidad-Fecha
nrd=nrow(datos)
e0=rep(NA,nrd)
for (i in 1:nrd){
  if (datos$ENT[i] < 10){
    e0[i]=paste0("0",datos$ENT[i])
  }
  else{
    e0[i]=datos$ENT[i]
  }
}

ID=paste(e0,datos$FECHA,sep="-")
datos_man=cbind.data.frame(ID,datos)
datos_man=rename(datos_man,MAN_EST=EST,MAN_TRA=TRA,MAN_SAL=SAL,MAN_SEG=SEG,MAN_PRE=PRE,MAN_HOR=HOR,MAN_EXP=EXP)


########## SEGMENTO NO MANUFACTURERO ##########
#Leer base
#nombre_base_csv=paste0("immex_nomanu_ent_2009_01_",fcsv,".csv")
#base_csv=read.csv(nombre_base_csv)
base_csv=base_csv_nomanu

#Crear variable de mes con ceros
nrb=nrow(base_csv)
m0=rep(NA,nrb)
for (i in 1:nrb){
  if (base_csv$ID_MES[i] < 10){
    m0[i]=paste0("0",base_csv$ID_MES[i])
  }
  else{
    m0[i]=base_csv$ID_MES[i]
  }
}

#Crear variable de fecha con anio y mes
FECHA=paste(base_csv$ANIO,m0,sep="-")

#Base de IMMEX No Manufacturero con fecha (variables seleccionadas)
varsec=cbind.data.frame(FECHA,base_csv)
varsec=select(varsec,-ANIO,-ID_MES,ENT=ID_ENTIDAD,EST=NUM_DE_ESTAB_ACT,-G210A,-M000B,-starts_with("IN"),-starts_with("K"),-ESTATUS)

#Variables de totales
#Total Trabajadores (Contratados y Subcontratados)
TRA=varsec$H010A+varsec$I000A
#Total Salarios (Contratados)
SAL=varsec$J100A_Y_J200A
#Total Horas Trabajadas (Contratados y Subcontratados)
HOR=varsec$H010D_Y_I000D
#Total Exportaciones (Ingresos por produccion, comercializacion y prestacion de servicios)
EXP=varsec$M000C

#Variables de Seguridad Social y Prestaciones (variables seleccionadas ademas de totales)
#Seguridad Social
SEG=varsec$J300A
#Prestaciones
PRE=varsec$J400A

#Base con totales y variables seleccionadas
base_tot=cbind.data.frame(varsec[1:3],TRA,SAL,SEG,PRE,HOR,EXP)


#Lista de entidades en el programa immex (no manufactureros)
e=unique(varsec$ENT)

#Colapsar datos de cada entidad por fecha
for (i in e){
  #Datos de Entidad
  datos_ent=base_tot[base_tot$ENT==i,]
  datos_ent=select(datos_ent,-ENT)
  #Colapsar por mes todos los municipios del estado
  colaps=aggregate(.~FECHA, datos_ent, sum)
  colaps=mutate(colaps,ENT=i)
  colaps=colaps[,c(9,1:8)]
  if (i==e[1]){
    datos=colaps
  }else{
    datos=rbind.data.frame(datos,colaps)
  }
}

#Identificador Entidad-Fecha
nrd=nrow(datos)
e0=rep(NA,nrd)
for (i in 1:nrd){
  if (datos$ENT[i] < 10){
    e0[i]=paste0("0",datos$ENT[i])
  }
  else{
    e0[i]=datos$ENT[i]
  }
}

ID=paste(e0,datos$FECHA,sep="-")
datos_nom=cbind.data.frame(ID,datos)
datos_nom=rename(datos_nom,NOM_EST=EST,NOM_TRA=TRA,NOM_SAL=SAL,NOM_SEG=SEG,NOM_PRE=PRE,NOM_HOR=HOR,NOM_EXP=EXP)


########## UNION SEGMENTO MANUFACTURERO Y NO MANUFACTURERO ##########
datmen=merge(select(datos_man,-FECHA,-ENT),select(datos_nom,-FECHA,-ENT), by="ID",all=TRUE)

#Crear variable de fecha con anio, mes y dia
FECHA=as.Date(paste(substr(datmen$ID,start=4,stop=11),"01",sep="-"))

#Crear variable de entidad
ENT=as.numeric((substr(datmen$ID,start=1,stop=2)))

#Agregar fecha y entidad
datmen=cbind.data.frame(FECHA,ENT,datmen)
datmen=datmen[,c(3,2,1,4:17)]
datmen[is.na(datmen)]=0

#Suma Totales
datmen=mutate(datmen,TOT_EST=MAN_EST+NOM_EST,TOT_TRA=MAN_TRA+NOM_TRA,TOT_SAL=MAN_SAL+NOM_SAL,TOT_SEG=MAN_SEG+NOM_SEG,TOT_PRE=MAN_PRE+NOM_PRE,TOT_HOR=MAN_HOR+NOM_HOR,TOT_EXP=MAN_EXP+NOM_EXP)


########## PROMEDIO DE DATOS MENSUALES DE LOS ULTIMOS 12 MESES ##########
#Lista de entidades en el programa immex (manufactureros y no manufactureros)
e=unique(datmen$ENT)
e=sort(e)
el=length(e)

#Crear base vacia
pro_datmen_12m=matrix(data=0,(nrow(datmen)-(12*el)),ncol(datmen))
pro_datmen_12m=as.data.frame(pro_datmen_12m)
names(pro_datmen_12m)=names(datmen)

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  #Datos de Entidad
  datos_ent=datmen[datmen$ENT==i,]
  #Promedio 12 meses
  for (j in 1:(nrow(datos_ent)-12)){
    #contador
    f=f+1
    pro_datmen_12m[f,1]=as.character(datos_ent[j+12,1])
    pro_datmen_12m[f,2]=i
    pro_datmen_12m[f,3]=as.character(datos_ent[j+12,3])
    for (k in 4:ncol(datos_ent)){
      if(!is.na(datos_ent[j,k])){
        pro_datmen_12m[f,k]=mean(datos_ent[(j+1):(j+12),k])
      }else{
        pro_datmen_12m[f,k]=NA
      }
    }
  }
}


########## VARIACIONES ANUALES DE DATOS MENSUALES (respecto al mismo periodo del mes anterior) ##########
#Lista de entidades en el programa immex (manufactureros y no manufactureros)
e=unique(datmen$ENT)
e=sort(e)
el=length(e)

#Crear base vacia
var_datmen_12m=matrix(data=0,(nrow(datmen)-(12*el)),ncol(datmen))
var_datmen_12m=as.data.frame(var_datmen_12m)
names(var_datmen_12m)=names(datmen)

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  #Datos de Entidad
  datos_ent=datmen[datmen$ENT==i,]
  #Variacion 12 meses
  for (j in 1:(nrow(datos_ent)-12)){
    #contador
    f=f+1
    var_datmen_12m[f,1]=as.character(datos_ent[j+12,1])
    var_datmen_12m[f,2]=i
    var_datmen_12m[f,3]=as.character(datos_ent[j+12,3])
    for (k in 4:ncol(datos_ent)){
      if(datos_ent[j,k]!=0){
        var_datmen_12m[f,k]=(datos_ent[j+12,k]/datos_ent[j,k]-1)*100
      }else{
        var_datmen_12m[f,k]=NA
      }
    }
  }
}


#Promedio de variaciones anuales de datos mensuales
#Crear base vacia
pro_var_datmen_12m=matrix(data=0,(nrow(var_datmen_12m)-(12*el)),ncol(var_datmen_12m))
pro_var_datmen_12m=as.data.frame(pro_var_datmen_12m)
names(pro_var_datmen_12m)=names(var_datmen_12m)

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  #Datos de Entidad
  datos_ent=var_datmen_12m[var_datmen_12m$ENT==i,]
  #Variacion 12 meses
  for (j in 1:(nrow(datos_ent)-12)){
    #contador
    f=f+1
    pro_var_datmen_12m[f,1]=as.character(datos_ent[j+12,1])
    pro_var_datmen_12m[f,2]=i
    pro_var_datmen_12m[f,3]=as.character(datos_ent[j+12,3])
    for (k in 4:ncol(datos_ent)){
      if(!is.na(datos_ent[j,k])){
        pro_var_datmen_12m[f,k]=mean(datos_ent[(j+1):(j+12),k])
      }else{
        pro_var_datmen_12m[f,k]=NA
      }
    }
  }
}


########## CIFRAS DE ACUMULADOS ANUALES Y SUS VARIACIONES ANUALES ##########
#(suma de ultimos 12 meses respecto al mismo periodo del mes anterior)
#Cifras de personal en promedio mensual de personas; cifras de remuneraciones y exportaciones en millones de pesos anuales; cifras de horas trabajadas en millones de horas anuales; cifras de establecimientos en numero de establecimientos

#Lista de entidades en el programa immex (manufactureros y no manufactureros)
e=unique(datmen$ENT)
e=sort(e)
el=length(e)

#Crear base vacia
datacu=matrix(data=0,(nrow(datmen)-(12*el)),ncol(datmen))
datacu=as.data.frame(datacu)
names(datacu)=names(datmen)

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  #Datos de Entidad
  datos_ent=datmen[datmen$ENT==i,]
  #Acumulado de 12 meses
  for (j in 1:(nrow(datos_ent)-12)){
    #contador
    f=f+1
    datacu[f,1]=as.character(datos_ent[j+12,1])
    datacu[f,2]=i
    datacu[f,3]=as.character(datos_ent[j+12,3])
    datacu[f,4]=datos_ent[j+12,4]
    datacu[f,5]=datos_ent[j+12,5]
    for (k in 6:10){
      datacu[f,k]=sum(datos_ent[(j+1):(j+12),k])/1000
    }
    datacu[f,11]=datos_ent[j+12,11]
    datacu[f,12]=datos_ent[j+12,12]
    for (k in 13:17){
      datacu[f,k]=sum(datos_ent[(j+1):(j+12),k])/1000
    }
    datacu[f,18]=datos_ent[j+12,18]
    datacu[f,19]=datos_ent[j+12,19]
    for (k in 20:24){
      datacu[f,k]=sum(datos_ent[(j+1):(j+12),k])/1000
    }
  }
}


#Variacionesde acumulados anuales
#Crear base vacia
var_acu=matrix(data=0,(nrow(datacu)-(12*el)),ncol(datacu))
var_acu=as.data.frame(var_acu)
names(var_acu)=names(datacu)

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  #Datos de Entidad
  datos_ent=datacu[datacu$ENT==i,]
  #Variacion 12 meses
  for (j in 1:(nrow(datos_ent)-12)){
    #contador
    f=f+1
    var_acu[f,1]=as.character(datos_ent[j+12,1])
    var_acu[f,2]=i
    var_acu[f,3]=as.character(datos_ent[j+12,3])
    for (k in 4:ncol(datos_ent)){
      if(datos_ent[j,k]!=0){
        var_acu[f,k]=(datos_ent[j+12,k]/datos_ent[j,k]-1)*100
      }else{
        var_acu[f,k]=NA
      }
    }
  }
}


#Promedio de variaciones de acumulados anuales
#Crear base vacia
pro_var_acu=matrix(data=0,(nrow(var_acu)-(12*el)),ncol(var_acu))
pro_var_acu=as.data.frame(pro_var_acu)
names(pro_var_acu)=names(var_acu)

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  #Datos de Entidad
  datos_ent=var_acu[var_acu$ENT==i,]
  #Variacion 12 meses
  for (j in 1:(nrow(datos_ent)-12)){
    #contador
    f=f+1
    pro_var_acu[f,1]=as.character(datos_ent[j+12,1])
    pro_var_acu[f,2]=i
    pro_var_acu[f,3]=as.character(datos_ent[j+12,3])
    for (k in 4:ncol(datos_ent)){
      if(!is.na(datos_ent[j,k])){
        pro_var_acu[f,k]=mean(datos_ent[(j+1):(j+12),k])
      }else{
        pro_var_acu[f,k]=NA
      }
    }
  }
}


########## DISTRIBUCIONES POR ENTIDAD ##########
#Nombres de Entidades
nent=c("Aguascalientes","Baja California","Coahuila","Chihuahua","Ciudad de México","Durango","Guanajuato","Jalisco","Estado de México","Michoacán","Nuevo León","Puebla","Querétaro","San Luis Potosí","Sinaloa","Sonora","Tamaulipas","Veracruz","Yucatán","Otras")
#nent=read.csv("Entidad_Federativa_IMMEX.csv")

#Seleccionar ultimo dato de cada entidad federativa

#Lista de entidades en el programa immex (manufactureros y no manufactureros)
e=unique(datmen$ENT)
e=sort(e)
el=length(e)

#Crear base vacia
ult_dat=matrix(data=0,el,ncol(datmen)-1)
ult_dat=as.data.frame(ult_dat)
names(ult_dat)=c(names(datmen)[2],"EF",names(datmen)[4:24])

#Contador de fila
f=0

#Datos por Entidad
for (i in e){
  f=f+1
  #Datos de Entidad
  datos_ent=datmen[datmen$ENT==i,]
  #datos_ent=mutate(datos_ent,EF=as.character(nent[f,1]))
  datos_ent=mutate(datos_ent,EF=as.character(nent[f]))
  datos_ent=datos_ent[,c(1:2,25,3:24)]
  datos_ent=select(datos_ent,-ID,-FECHA)
  #Seleccionar ultimo dato de cada entidad
  ult_dat[f,]=datos_ent[nrow(datos_ent),]
}


#Crear base vacia
dis_ult_dat=matrix(data=0,el,ncol(ult_dat))
dis_ult_dat=as.data.frame(dis_ult_dat)
names(dis_ult_dat)=names(ult_dat)

#Sumas de columnas
sumcol=lapply(ult_dat[,3:ncol(ult_dat)],sum)
sumcol=as.numeric(sumcol)

#Datos por fila
for (i in 1:el){
  dis_ult_dat[i,1]=ult_dat[i,1]
  dis_ult_dat[i,2]=ult_dat[i,2]
  for (j in 3:ncol(ult_dat)){
    dis_ult_dat[i,j]=(ult_dat[i,j]/sumcol[j-2])*100
  }
}


########## BASES PARA BOLETIN ##########
#Datos Mensuales de Jalisco en Millones de Pesos y Millones de Horas desde 2013
jal_datmen_mdp=datmen[datmen$ENT==14,]
jal_datmen_mdp=jal_datmen_mdp[(nrow(jal_datmen_mdp)-(c+5*12)):nrow(jal_datmen_mdp),3:ncol(jal_datmen_mdp)]
for (i in 4:8){
  jal_datmen_mdp[,i]=jal_datmen_mdp[,i]/1000
}
for (i in 11:15){
  jal_datmen_mdp[,i]=jal_datmen_mdp[,i]/1000
}
for (i in 18:22){
  jal_datmen_mdp[,i]=jal_datmen_mdp[,i]/1000
}

#Promedio de Datos Mensuales de Jalisco en Millones de Pesos y Millones de Horas desde 2013
jal_datmen_mdp_pro=pro_datmen_12m[pro_datmen_12m$ENT==14,]
jal_datmen_mdp_pro=jal_datmen_mdp_pro[(nrow(jal_datmen_mdp_pro)-(c+5*12)):nrow(jal_datmen_mdp_pro),3:ncol(jal_datmen_mdp_pro)]
for (i in 4:8){
  jal_datmen_mdp_pro[,i]=jal_datmen_mdp_pro[,i]/1000
}
for (i in 11:15){
  jal_datmen_mdp_pro[,i]=jal_datmen_mdp_pro[,i]/1000
}
for (i in 18:22){
  jal_datmen_mdp_pro[,i]=jal_datmen_mdp_pro[,i]/1000
}

#Variaciones de Datos Mensuales de Jalisco (ultimos 12 meses) desde 2013
jal_datmen_var=var_datmen_12m[var_datmen_12m$ENT==14,]
jal_datmen_var=jal_datmen_var[(nrow(jal_datmen_var)-(c+5*12)):nrow(jal_datmen_var),3:ncol(jal_datmen_var)]

#Promedio de Variaciones de Datos Mensuales de Jalisco (ultimos 12 meses) desde 2013
jal_datmen_var_pro=pro_var_datmen_12m[pro_var_datmen_12m$ENT==14,]
jal_datmen_var_pro=jal_datmen_var_pro[(nrow(jal_datmen_var_pro)-(c+5*12)):nrow(jal_datmen_var_pro),3:ncol(jal_datmen_var_pro)]

#Datos de Acumulados Anuales para Jalisco desde 2013
jal_datacu=datacu[datacu$ENT==14,]
jal_datacu=jal_datacu[(nrow(jal_datacu)-(c+5*12)):nrow(jal_datacu),3:ncol(jal_datacu)]

#Variaciones de Datos de Acumulados Anuales para Jalisco desde 2013
jal_datacu_var=var_acu[var_acu$ENT==14,]
jal_datacu_var=jal_datacu_var[(nrow(jal_datacu_var)-(c+5*12)):nrow(jal_datacu_var),3:ncol(jal_datacu_var)]

#Promedio de Variaciones de Datos de Acumulados Anuales para Jalisco desde 2013
jal_datacu_var_pro=pro_var_acu[pro_var_acu$ENT==14,]
jal_datacu_var_pro=jal_datacu_var_pro[(nrow(jal_datacu_var_pro)-(c+5*12)):nrow(jal_datacu_var_pro),3:ncol(jal_datacu_var_pro)]


########## DATOS PARA CADA GRAFICAS ##########
#Fechas para graficar
#Crear base vacia
fecha_gra=matrix(data="",nrow(jal_datmen_mdp),3)
fecha_gra=as.data.frame(fecha_gra)
fecha_gra <- data.frame(lapply(fecha_gra, as.character), stringsAsFactors=FALSE)
names(fecha_gra)=c("FECHA","ANIO","MES")

fecha_gra$FECHA=jal_datmen_mdp$FECHA

for (i in 1:nrow(fecha_gra)){
  if (substr(fecha_gra$FECHA[i],6,7)=="01"){
    fecha_gra$ANIO[i]=substr(fecha_gra$FECHA[i],1,4)
    fecha_gra$MES[i]="ENE"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="02"){
    fecha_gra$MES[i]="FEB"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="03"){
    fecha_gra$MES[i]="MAR"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="04"){
    fecha_gra$MES[i]="ABR"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="05"){
    fecha_gra$MES[i]="MAY"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="06"){
    fecha_gra$MES[i]="JUN"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="07"){
    fecha_gra$MES[i]="JUL"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="08"){
    fecha_gra$MES[i]="AGO"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="09"){
    fecha_gra$MES[i]="SEP"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="10"){
    fecha_gra$MES[i]="OCT"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="11"){
    fecha_gra$MES[i]="NOV"
  }
  if (substr(fecha_gra$FECHA[i],6,7)=="12"){
    fecha_gra$MES[i]="DIC"
  }
}


#Manufacturas: Exportaciones Mensuales con su Promedio
man_exp_men_pro=cbind.data.frame(fecha_gra[,c(2,3)],jal_datmen_mdp[,8],jal_datmen_mdp_pro[,8])
names(man_exp_men_pro)=c("Anio","Mes","Exportaciones","Promedio")
man_exp_men_pro$Exportaciones=round(man_exp_men_pro$Exportaciones)
man_exp_men_pro$Promedio=round(man_exp_men_pro$Promedio)
#Tabla
man_exp_men_pro_tab=cbind.data.frame(man_exp_men_pro[nrow(man_exp_men_pro),3],man_exp_men_pro[nrow(man_exp_men_pro)-12,3],round(jal_datmen_var[nrow(jal_datmen_var),8],1),round(jal_datmen_var[nrow(jal_datmen_var)-1,8],1),man_exp_men_pro[nrow(man_exp_men_pro),4],man_exp_men_pro[nrow(man_exp_men_pro)-1,4])
names(man_exp_men_pro_tab)=c(as.character(jal_datmen_mdp[nrow(jal_datmen_mdp),1]),as.character(jal_datmen_mdp[nrow(jal_datmen_mdp)-12,1]),paste(man_exp_men_pro[nrow(man_exp_men_pro),2],"%"),paste(man_exp_men_pro[nrow(man_exp_men_pro)-1,2],"%"),paste(man_exp_men_pro[nrow(man_exp_men_pro),2],"PROM"),paste(man_exp_men_pro[nrow(man_exp_men_pro)-1,2],"PROM"))


#No Manufacturas: Exportaciones Mensuales con su Promedio
nom_exp_men_pro=cbind.data.frame(fecha_gra[,c(2,3)],jal_datmen_mdp[,15],jal_datmen_mdp_pro[,15])
names(nom_exp_men_pro)=c("Anio","Mes","Exportaciones","Promedio")
nom_exp_men_pro$Exportaciones=round(nom_exp_men_pro$Exportaciones)
nom_exp_men_pro$Promedio=round(nom_exp_men_pro$Promedio)
#Tabla
nom_exp_men_pro_tab=cbind.data.frame(nom_exp_men_pro[nrow(nom_exp_men_pro),3],nom_exp_men_pro[nrow(nom_exp_men_pro)-12,3],round(jal_datmen_var[nrow(jal_datmen_var),15],1),round(jal_datmen_var[nrow(jal_datmen_var)-1,15],1),nom_exp_men_pro[nrow(nom_exp_men_pro),4],nom_exp_men_pro[nrow(nom_exp_men_pro)-1,4])
names(nom_exp_men_pro_tab)=c(as.character(jal_datmen_mdp[nrow(jal_datmen_mdp),1]),as.character(jal_datmen_mdp[nrow(jal_datmen_mdp)-12,1]),paste(nom_exp_men_pro[nrow(nom_exp_men_pro),2],"%"),paste(nom_exp_men_pro[nrow(nom_exp_men_pro)-1,2],"%"),paste(nom_exp_men_pro[nrow(nom_exp_men_pro),2],"PROM"),paste(nom_exp_men_pro[nrow(nom_exp_men_pro)-1,2],"PROM"))


#Exportaciones Totales Acumulado Anual
tot_exp_acu=cbind.data.frame(fecha_gra[,c(2,3)],jal_datacu[,22],jal_datacu_var[,22],jal_datacu_var_pro[,22])
names(tot_exp_acu)=c("Anio","Mes","Exportaciones","Variacion","Variacion promedio")
tot_exp_acu$Exportaciones=round(tot_exp_acu$Exportaciones)
tot_exp_acu$Variacion=round(tot_exp_acu$Variacion,1)
tot_exp_acu$`Variacion promedio`=round(tot_exp_acu$`Variacion promedio`,1)
#Tabla
tot_exp_acu_tab=cbind.data.frame(tot_exp_acu[nrow(tot_exp_acu),3],tot_exp_acu[nrow(tot_exp_acu)-1,3],tot_exp_acu[nrow(tot_exp_acu),4],tot_exp_acu[nrow(tot_exp_acu)-1,4],tot_exp_acu[nrow(tot_exp_acu),5])
names(tot_exp_acu_tab)=c(paste(tot_exp_acu[nrow(tot_exp_acu),2],"$"),paste(tot_exp_acu[nrow(tot_exp_acu)-1,2],"$"),paste(tot_exp_acu[nrow(tot_exp_acu),2],"%"),paste(tot_exp_acu[nrow(tot_exp_acu)-1,2],"%"),"PROM %")


#Establecimientos Totales
tot_est=cbind.data.frame(fecha_gra[,c(2,3)],jal_datmen_mdp[,16],jal_datmen_mdp_pro[,16])
names(tot_est)=c("Anio","Mes","Establecimientos","Promedio")
tot_est$Promedio=round(tot_est$Promedio)
#Tabla
tot_est_tab=cbind.data.frame(tot_est[nrow(tot_est),3],tot_est[nrow(tot_est)-1,3],tot_est[nrow(tot_est)-12,3],round(jal_datmen_var[nrow(jal_datmen_var),16],1),tot_est[nrow(tot_est),4],tot_est[nrow(tot_est)-1,4])
names(tot_est_tab)=c(paste(tot_est[nrow(tot_est),2],"EST"),paste(tot_est[nrow(tot_est)-1,2],"EST"),paste(jal_datmen_mdp[nrow(jal_datmen_mdp)-12,1],"EST"),"% 12M",paste(tot_est[nrow(tot_est),2],"PROM"),paste(tot_est[nrow(tot_est)-1,2],"PROM"))


#Trabajadores Totales
tot_tra=cbind.data.frame(fecha_gra[,c(2,3)],jal_datmen_mdp[,17],jal_datmen_mdp_pro[,17])
names(tot_tra)=c("Anio","Mes","Trabajadores","Promedio")
tot_tra$Trabajadores=round(tot_tra$Trabajadores)
tot_tra$Promedio=round(tot_tra$Promedio)
#Tabla: NOV, OCT, 2017-11-01, NOV%, OCT%, NOV PROM, OCT PROM (7)
tot_tra_tab=cbind.data.frame(tot_tra[nrow(tot_tra),3],tot_tra[nrow(tot_tra)-1,3],tot_tra[nrow(tot_tra)-12,3],round(jal_datmen_var[nrow(jal_datmen_var),17],1),round(jal_datmen_var[nrow(jal_datmen_var)-1,17],1),tot_tra[nrow(tot_tra),4],tot_tra[nrow(tot_tra)-1,4])
names(tot_tra_tab)=c(tot_tra[nrow(tot_tra),2],tot_tra[nrow(tot_tra)-1,2],jal_datmen_var[nrow(jal_datmen_var)-12,1],paste(tot_tra[nrow(tot_tra),2],"%"),paste(tot_tra[nrow(tot_tra)-1,2],"%"),paste(tot_tra[nrow(tot_tra),2],"PROM"),paste(tot_tra[nrow(tot_tra)-1,2],"PROM"))


#Variaciones Respecto al Mismo Mes del Anio Anterior de Trabajadores Totales
var_tot_tra=cbind.data.frame(fecha_gra[,c(2,3)],jal_datmen_var[,17],jal_datmen_var_pro[,17])
names(var_tot_tra)=c("Anio","Mes","Variacion","Variacion promedio")
var_tot_tra$Variacion=round(var_tot_tra$Variacion,1)
var_tot_tra$`Variacion promedio`=round(var_tot_tra$`Variacion promedio`,1)
#Tabla: NOV, OCT, 2017-11-01, NOV PROM, OCT PROM
var_tot_tra_tab=cbind.data.frame(var_tot_tra[nrow(var_tot_tra),3],var_tot_tra[nrow(var_tot_tra)-1,3],var_tot_tra[nrow(var_tot_tra)-12,3],var_tot_tra[nrow(var_tot_tra),4],var_tot_tra[nrow(var_tot_tra)-1,4])
names(var_tot_tra_tab)=c(var_tot_tra[nrow(var_tot_tra),2],var_tot_tra[nrow(var_tot_tra)-1,2],jal_datmen_var[nrow(jal_datmen_var)-12,1],paste(var_tot_tra[nrow(var_tot_tra),2],"PROM"),paste(var_tot_tra[nrow(var_tot_tra)-1,2],"PROM"))


#Distribucion Porcentual de los Establecimientos Manufactureros y No Manufactureros con programa IMMEX por Entidad Federativa
dis_tot_est=cbind.data.frame(dis_ult_dat$EF,dis_ult_dat$TOT_EST)
names(dis_tot_est)=c("Entidad","Distribucion")
#dis_tot_est$Distribucion=round(dis_tot_est$Distribucion,1)
#Quitar "Otras"
dis_tot_est=dis_tot_est[1:(nrow(dis_tot_est)-1),]
dis_tot_est=dis_tot_est[order(dis_tot_est$Distribucion),]


#Distribucion Porcentual de Personal Ocupado en los Establecimientos Manufactureros y No Manufactureros con programa IMMEX por Entidad Federativa
dis_tot_tra=cbind.data.frame(dis_ult_dat$EF,dis_ult_dat$TOT_TRA)
names(dis_tot_tra)=c("Entidad","Distribucion")
#dis_tot_tra$Distribucion=round(dis_tot_tra$Distribucion,1)
#Quitar "Otras"
dis_tot_tra=dis_tot_tra[1:(nrow(dis_tot_tra)-1),]
dis_tot_tra=dis_tot_tra[order(dis_tot_tra$Distribucion),]


#Borrar variables y bases que ya no se necesitan
rm(list=c("base_csv","base_tot","colaps","datos","datos_ent","datos_man","datos_nom","dis_ult_dat","e","el","e0","ENT","EXP","f","FECHA","fecha_gra","HOR","i","ID","j","k","m0","nent","nrb","nrd","PRE","SAL","SEG","sumcol","TRA","ult_dat","varsec"))


########## TEXTO DESCRIPCIÓN 1: man_exp_men_pro ##########
#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
ud = man_exp_men_pro[nrow(man_exp_men_pro),3] #último dato
ad = man_exp_men_pro[nrow(man_exp_men_pro)-12,3] #año anterior dato
uv = round(jal_datmen_var[nrow(jal_datmen_var),8],1) #última variación
av = round(jal_datmen_var[nrow(jal_datmen_var)-1,8],1) #año anterior variación
up = man_exp_men_pro[nrow(man_exp_men_pro),4] #último promedio
ap = man_exp_men_pro[nrow(man_exp_men_pro)-1,4] #mes anterior promedio

descripcion1_1 = paste0("La estadística mensual sobre establecimientos con programa de la Industria Manufacturera, Maquiladora y de Servicios de Exportación (IMMEX) es publicada por INEGI de manera mensual. En ",
                        mesactual, " de ",añoactual," se difundió el dato correspondiente a ", mesp, " de ")
if (mesactual == "enero" | mesactual == "febrero"){
  descripcion1_1 = paste0(descripcion1_1,year(ymd(Sys.Date()))-1,".")
} else {
  descripcion1_1 = paste0(descripcion1_1,year(ymd(Sys.Date())),".")
}


descripcion1_2 = paste0("De acuerdo con las cifras reportadas, los ingresos provenientes del mercado extranjero que obtuvieron los establecimientos manufactureros en el programa IMMEX en Jalisco se ubicaron en ",
                        format(ud,big.mark=","), " millones de pesos durante el mes de ", mesp, " de ", año, ", ")

if(ud > ad){
  descripcion1_2 = paste0(descripcion1_2,"por encima de la cifra de ",mesp, " de ", año-1,", la cual se ubicaba en ",
                          format(ad,big.mark=",")," millones de pesos. ")
}else if(ud < ad){
  descripcion1_2 = paste0(descripcion1_2,"por debajo de la cifra de ",mesp, " de ", año-1,", la cual se ubicaba en ",
                          format(ad,big.mark=",")," millones de pesos. ")
}else{
  descripcion1_2 = paste0(descripcion1_2,"en línea con la cifra de ",mesp, " de ", año-1)
}


if(uv > 0){
  descripcion1_2 = paste0(descripcion1_2,"Esto representó un crecimiento anual de ",format(uv,nsmall = 1),"%, ")
  if(uv > av & av >= 0){
    descripcion1_2 = paste0(descripcion1_2,"incremento superior al observado el mes anterior cuando se presentó un aumento de los ingresos de ",
                            format(abs(av),nsmall = 1),"% anual. ")
  }else if(uv < av){
    descripcion1_2 = paste0(descripcion1_2,"incremento inferior al observado el mes anterior cuando se presentó un aumento de los ingresos de ",
                            format(abs(av),nsmall = 1),"% anual. ")
  }else if(uv > av & av < 0){
    descripcion1_2 = paste0(descripcion1_2,"cifra superior a la observada el mes anterior cuando se presentó una disminución de los ingresos de ",
                            format(abs(av),nsmall = 1),"% anual. ")
  }
}else if(uv < 0){
  descripcion1_2 = paste0(descripcion1_2,"Esto representó una disminución anual de ",format(abs(uv),nsmall = 1),"%, ")
  if(uv > av){
    descripcion1_2 = paste0(descripcion1_2,"caída menor a la observada el mes anterior cuando se presentó un descenso de los ingresos de ",
                            format(abs(av),nsmall = 1),"% anual. ")
  }else if(uv < av & av < 0){
    descripcion1_2 = paste0(descripcion1_2,"caída mayor a la observada el mes anterior cuando se presentó un descenso de los ingresos de ",
                            format(abs(av),nsmall = 1),"% anual. ")
  }else if(uv < av & av >= 0){
    descripcion1_2 = paste0(descripcion1_2,"cifra inferior a la observada el mes anterior cuando se presentó un crecimiento de los ingresos de ",
                            format(abs(av),nsmall = 1),"% anual. ")
  }
}

#Conector
if(uv > av){
  if(up > ap){conector1 = "Además, "}
  else{conector1 = "Sin embargo, "}
}else{
  if(up > ap){conector1 = "Sin embargo, "}
  else{conector1 = "Además, "}
}
descripcion1_2 = paste0(descripcion1_2,conector1,"el promedio de los últimos doce meses ")

#Promedio
if(up > ap){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(man_exp_men_pro[nrow(man_exp_men_pro)-ac,4] > (man_exp_men_pro[nrow(man_exp_men_pro)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(man_exp_men_pro[nrow(man_exp_men_pro)-ac-1-dcp,4] < man_exp_men_pro[nrow(man_exp_men_pro)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion1_2 = paste0(descripcion1_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion1_2 = paste0(descripcion1_2,"registró su primer aumento después de ",dcp,
                            " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion1_2 = paste0(descripcion1_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion1_2 = paste0(descripcion1_2,"mantiene su tendencia creciente")
  }
  descripcion1_2 = paste0(descripcion1_2,", al pasar de ",format(ap,big.mark=","), " millones a ",format(up,big.mark=","), " millones en ",mesp, " de ",
                          año, " respecto al mes inmediato anterior.")
}else if( up < ap){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(man_exp_men_pro[nrow(man_exp_men_pro)-dc,4] < (man_exp_men_pro[nrow(man_exp_men_pro)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(man_exp_men_pro[nrow(man_exp_men_pro)-dc-1-acp,4] > man_exp_men_pro[nrow(man_exp_men_pro)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion1_2 = paste0(descripcion1_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion1_2 = paste0(descripcion1_2,"registró su primer descenso después de ",acp,
                            " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion1_2 = paste0(descripcion1_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion1_2 = paste0(descripcion1_2,"mantiene su tendencia decreciente")
  }
  descripcion1_2 = paste0(descripcion1_2,", al pasar de ",format(ap,big.mark=","), " millones a ",format(up,big.mark=","), " millones en ",mesp, " de ",
                          año, " respecto al mes inmediato anterior.")
}else{
  descripcion1_2 = paste0(descripcion1_2,conector1,"se mantuvo sin cambios.")
}


########## TEXTO DESCRIPCIÓN 2: nom_exp_men_pro ##########
#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
ud = nom_exp_men_pro[nrow(nom_exp_men_pro),3] #último dato
ad = nom_exp_men_pro[nrow(nom_exp_men_pro)-12,3] #año anterior dato
uv = round(jal_datmen_var[nrow(jal_datmen_var),15],1) #última variación
av = round(jal_datmen_var[nrow(jal_datmen_var)-1,15],1) #año anterior variación
up = nom_exp_men_pro[nrow(nom_exp_men_pro),4] #último promedio
ap = nom_exp_men_pro[nrow(nom_exp_men_pro)-1,4] #mes anterior promedio

descripcion_2 <- paste0("Por otro lado, los ingresos provenientes del mercado extranjero que obtuvieron los establecimientos",
                        " no manufactureros en el programa IMMEX en Jalisco se ubicaron en ",
                        format(ud, big.mark = ","), " millones de pesos durante el mes de ",
                        mesp, " de ", año, ",")

if(ud > ad){
  descripcion_2 <- paste0(descripcion_2, " cifra superior a la de ", mesp, " de ",
                          año-1, ", la cual alcanzó los ", format(ad, big.mark = ","),
                          " millones de pesos. ")
} else if(ud < ad){
  descripcion_2 <- paste0(descripcion_2, " cifra inferior a la de ", mesp, " de ",
                          año-1, ", la cual alcanzó los ", format(ad, big.mark = ","),
                          " millones de pesos. ")
} else {
  descripcion_2 <- paste0(descripcion_2, " en linea con la cifra de ", mesp, " de ",
                          año-1, ". ")
}

if(uv > 0){
  descripcion_2 <- paste0(descripcion_2, "Esto representó un crecimiento anual de ",
                          format(abs(uv),nsmall = 1), "%")
  if(av == 0){
    descripcion_2 <- paste0(descripcion_2, ". ")
  }else if(uv > av & av > 0){
    descripcion_2 <- paste0(descripcion_2, ", incremento superior al observado el mes anterior cuando se presentó un aumento de los ingresos de ",
                            format(abs(av),nsmall = 1), "% anual. ")
  }else if(uv < av){
    descripcion_2 <- paste0(descripcion_2, ", incremento inferior al observado el mes anterior cuando se presentó un aumento de los ingresos de ",
                            format(abs(av),nsmall = 1), "% anual. ")
  }else if(uv > av & av < 0){
    descripcion_2 <- paste0(descripcion_2, ", cifra superior a la observada el mes anterior cuando se presentó un descenso de los ingresos de ",
                            format(abs(av),nsmall = 1), "% anual. ")
  }
}else if(uv < 0){
  descripcion_2 <- paste0(descripcion_2, "Esto representó una disminución anual de ",
                          format(abs(uv),nsmall = 1), "%")
  if(av==0){
    descripcion_2 <- paste0(descripcion_2, ". ")
  }else if(uv > av){
    descripcion_2 <- paste0(descripcion_2, ", caída menor a la observada el mes anterior cuando se presentó un descenso de los ingresos de ",
                            format(abs(av),nsmall = 1), "% anual. ")
  }else if(uv < av & av < 0){
    descripcion_2 <- paste0(descripcion_2, ", caída mayor a la observada el mes anterior cuando se presentó un descenso de los ingresos de ",
                            format(abs(av),nsmall = 1), "% anual. ")
  }else if(uv < av & av > 0){
    descripcion_2 <- paste0(descripcion_2, ", cifra inferior a la observada el mes anterior cuando se presentó un crecimiento de los ingresos de ",
                            format(abs(av),nsmall = 1), "% anual. ")
  }
}

#Conector
if(uv > av){
  if(up > ap){
    descripcion_2 <- paste0(descripcion_2, "Además")
  }else{
    descripcion_2 <- paste0(descripcion_2, "Sin embargo")
  }
}else{
  if(up > ap){
    descripcion_2 <- paste0(descripcion_2, "Sin embargo")
  }else{
    descripcion_2 <- paste0(descripcion_2, "Además")
  }
}

descripcion_2 <- paste0(descripcion_2, ", el promedio de los últimos doce meses ")

rm(ac, dcp, dc, acp)

#Promedio
if(up > ap){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(nom_exp_men_pro[nrow(nom_exp_men_pro)-ac,4] > (nom_exp_men_pro[nrow(nom_exp_men_pro)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(nom_exp_men_pro[nrow(nom_exp_men_pro)-ac-1-dcp,4] < nom_exp_men_pro[nrow(nom_exp_men_pro)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_2 = paste0(descripcion_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion_2 = paste0(descripcion_2,"registró su primer aumento después de ",dcp,
                            " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion_2 = paste0(descripcion_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion_2 = paste0(descripcion_2,"mantiene su tendencia creciente")
  }
  descripcion_2 = paste0(descripcion_2,", al pasar de ",format(ap,big.mark=","), " millones a ",format(up,big.mark=","), " millones en ",mesp, " de ",
                          año, " respecto al mes inmediato anterior.")
}else if( up < ap){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(nom_exp_men_pro[nrow(nom_exp_men_pro)-dc,4] < (nom_exp_men_pro[nrow(nom_exp_men_pro)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(nom_exp_men_pro[nrow(nom_exp_men_pro)-dc-1-acp,4] > nom_exp_men_pro[nrow(nom_exp_men_pro)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_2 = paste0(descripcion_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion_2 = paste0(descripcion_2,"registró su primer descenso después de ",acp,
                            " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion_2 = paste0(descripcion_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion_2 = paste0(descripcion_2,"mantiene su tendencia decreciente")
  }
  descripcion_2 = paste0(descripcion_2,", al pasar de ",format(ap,big.mark=","), " millones a ",format(up,big.mark=","), " millones en ",mesp, " de ",
                          año, " respecto al mes inmediato anterior.")
}else{
  descripcion_2 = paste0(descripcion_2,"se mantuvo sin cambios.")
}



# TEXTO DESCRIPCIÓN 3: tot_exp_acu ----------------------------------------

#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
ud = tot_exp_acu[nrow(tot_exp_acu),3] #último dato
ad = tot_exp_acu[nrow(tot_exp_acu)-1,3] #mes anterior dato
uv = round(tot_exp_acu[nrow(tot_exp_acu),4],1) #última variación
av = round(tot_exp_acu[nrow(tot_exp_acu)-1,4],1) #mes anterior variación
uvp = round(tot_exp_acu[nrow(tot_exp_acu),5],1) #última variación promedio


descripcion_3 <- paste0("Los ingresos acumulados de los últimos doce meses que obtuvieron el total ",
                        "de los establecimientos manufactureros y no manufactureros en el programa ",
                        "IMMEX en Jalisco se ubicaron en ", format(ud, big.mark = ","), " millones ",
                        "de pesos al mes de ", mesp, " de ", año, ", ")
  
if(ud > ad){
  descripcion_3 <- paste0(descripcion_3, "cifra superior a la del mes inmediato anterior la cual alcanzaba los ",
                          format(ad, big.mark = ","), 
                          " millones de pesos acumulados. ")
} else if(ud < ad){
  descripcion_3 <- paste0(descripcion_3, "cifra inferior a la del mes inmediato anterior la cual alcanzaba los ",
                          format(ad, big.mark = ","), 
                          " millones de pesos acumulados. ")
}else {
  descripcion_3 <- paste0(descripcion_3, "en linea con la cifra del mes inmediato anterior")
}

#conector:

if(uv > 0){
  if(ud > ad){
    descripcion_3 <- paste0(descripcion_3, "Además, ")
  }else if(ud < ad | ud == ad){
    descripcion_3 <- paste0(descripcion_3, "Sin embargo, ")
  }
}else {
  if(ud > ad){
    descripcion_3 <- paste0(descripcion_3, "Sin embargo, ")
  }else if(ud < ad | ud == ad){
    descripcion_3 <- paste0(descripcion_3, "Además, ")
  }
}

if(uv > 0){
  if(uv > av & av > 0){
    descripcion_3 <- paste0(descripcion_3, "se presentó un crecimiento anual de ",
                            format(uv, nsmall = 1), "%, incremento superior al ",
                            "observado el mes anterior cuando se presentó un crecimiento de ",
                            format(av, nsmall = 1), "% anual. ")
  }else if(uv > av & av < 0){
    descripcion_3 <- paste0(descripcion_3, "se presentó un crecimiento anual de ",
                            format(uv, nsmall = 1), "%, cifra superior a ",
                            "la observada el mes anterior cuando se presentó una reducción de ",
                            format(abs(av), nsmall = 1), "% anual. ")
  }else if(uv < av){
    descripcion_3 <- paste0(descripcion_3, "se presentó un crecimiento anual de ",
                            format(uv, nsmall = 1), "%, incremento inferior al ",
                            "observado el mes anterior cuando se presentó un crecimiento de ",
                            format(av, nsmall = 1), "% anual. ")
  }else if(uv == av){
    descripcion_3 <- paste0(descripcion_3, "se presentó un crecimiento anual de ",
                            format(uv, nsmall = 1), "%, en línea con el incremento ",
                            "observado el mes anterior. ")
  }
}else if(uv < 0){
  if(uv < av & av < 0){
    descripcion_3 <- paste0(descripcion_3, "se presentó una reducción anual de ",
                            format(abs(uv), nsmall = 1), "%, caída mayor a la observada ",
                            "el mes anterior cuando se presentó una reducción de ",
                            format(abs(av), nsmall = 1), "% anual. ")
  }else if(uv < av & av > 0){
    descripcion_3 <- paste0(descripcion_3, "se presentó una reducción anual de ",
                            format(abs(uv), nsmall = 1), "%, cifra inferior a la observada ",
                            "el mes anterior cuando se presentó un crecimiento de ",
                            format(av, nsmall = 1), "% anual. ")
  }else if(uv > av){
    descripcion_3 <- paste0(descripcion_3, "se presentó una reducción anual de ",
                            format(abs(uv), nsmall = 1), "%, caída menor a la observada ",
                            "el mes anterior cuando se presentó una reducción de ",
                            format(abs(av), nsmall = 1), "% anual. ")
  }else if(uv == av){
    descripcion_3 <- paste0(descripcion_3, "se presentó una reducción anual de ",
                            format(abs(uv), nsmall = 1), "%, en línea con la caída ",
                            "observada el mes anterior. ")
  }
}

descripcion_3 <- paste0(descripcion_3, "Esta variación se ubicó ")

if(uv > uvp){
  descripcion_3 <- paste0(descripcion_3, "por encima del promedio de las cifras ",
                          "anuales de los últimos doce meses, que se ubica en ",
                          format(uvp, nsmall = 1), "%.")
}else if(uv < uvp){
  descripcion_3 <- paste0(descripcion_3, "por debajo del promedio de las cifras ",
                          "anuales de los últimos doce meses, que se ubica en ",
                          format(uvp, nsmall = 1), "%.")
}else{
  descripcion_3 <- paste0(descripcion_3, "en línea con el promedio de las cifras ",
                          "anuales de los últimos doce meses.")
}



# TEXTO DESCRIPCIÓN 4: tot_est --------------------------------------------

#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
ud = tot_est[nrow(tot_est),3] #último dato
adm = tot_est[nrow(tot_est)-1,3] #mes anterior dato
ad = tot_est[nrow(tot_est)-12,3] #año anterior dato
uv =  round((ud-ad)/ad*100, 1) #última variación año
up = tot_est[nrow(tot_est), 4] #último promedio
ap = tot_est[nrow(tot_est)-1, 4] #mes anterior promedio


descripcion_4 <- "El número total de establecimientos manufactureros y no manufactureros con programa IMMEX en Jalisco "

if(ud > adm){
  descripcion_4 <- paste0(descripcion_4, "aumentó de ", adm, " a ", ud, " establecimientos en ",
                          mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else if(ud < adm){
  descripcion_4 <- paste0(descripcion_4, "disminuyó de ", adm, " a ", ud, " establecimientos en ",
                          mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else {
  descripcion_4 <- paste0(descripcion_4, "fue de ", ud, ", en línea con la cifra del mes inmediato anterior. ")
}


#Conector
if(ud > adm){
  if(ud > ad){
    descripcion_4 <- paste0(descripcion_4, "Además, ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "Sin embargo, ")
  }
}else{
  if(ud < ad | ud == ad){
    descripcion_4 <- paste0(descripcion_4, "Además, ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "Sin embargo, ")
  }
}


if(ud > ad){
  descripcion_4 <- paste0(descripcion_4, "la última cifra de establecimientos reportada es ",
                          "superior a la de ", mesp, " de ", año-1, ", la cual se ubicaba en ",
                          ad, " establecimientos, lo que representó un crecimiento anual de ",
                          format(abs(uv),nsmall = 1), "%. ")
}else if(ud < ad){
  descripcion_4 <- paste0(descripcion_4, "la última cifra de establecimientos reportada es ",
                          "inferior a la de ", mesp, " de ", año-1, ", la cual se ubicaba en ",
                          ad, " establecimientos, lo que representó un crecimiento anual de ",
                          format(abs(uv),nsmall = 1), "%. ")
}else{
  descripcion_4 <- paste0(descripcion_4, "la última cifra de establecimientos reportada ",
                          "se mantiene en línea con la de ", mesp, " de ", año-1, ". ")
}


#Conector2
if(ud > ad){
  if(up > ap){
    descripcion_4 <- paste0(descripcion_4, "Asimismo, ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "Por otro lado, ")
  }
}else{
  if(up < ap | up == ap){
    descripcion_4 <- paste0(descripcion_4, "Asimismo, ")
  }else{
    descripcion_4 <- paste0(descripcion_4, "Por otro lado, ")
  }
}


descripcion_4 <- paste0(descripcion_4, "el promedio de los últimos doce meses ")


if(up > ap){
  descripcion_4 <- paste0(descripcion_4, "aumentó de ", ap, " a ", up,
                          " establecimientos en ", mesp, " de ", año,
                          " respecto al mes inmediato anterior, ",
                          "con lo que la cifra más reciente de establecimientos ",
                          "se ubica ")
}else if(up < ap){
  descripcion_4 <- paste0(descripcion_4, "disminuyó de ", ap, " a ", up,
                          " establecimientos en ", mesp, " de ", año,
                          " respecto al mes inmediato anterior, ",
                          "con lo que la cifra más reciente de establecimientos ",
                          "se ubica " )
}else {
  descripcion_4 <- paste0(descripcion_4, "se mantiene en línea respecto al del mes inmediato anterior ",
                          "con lo que la cifra más reciente de establecimientos ",
                          "se ubica " )
}


if(ud > up){
  descripcion_4 <- paste0(descripcion_4, "por arriba del promedio.")
}else if(ud < up){
  descripcion_4 <- paste0(descripcion_4, "por debajo del promedio.")
}else{
  descripcion_4 <- paste0(descripcion_4, "en línea con el promedio.")
}


# TEXTO DESCRIPCIÓN 5: dis_tot_est ----------------------------------------

#Dataframes
re = mutate(arrange(dis_tot_est, desc(Distribucion)), n = 1:nrow(dis_tot_est), num = num_cardinales[1:nrow(dis_tot_est)])#Ranking estados
nf = filter(dis_tot_est, Entidad != "Baja California" , Entidad != "Sonora" , Entidad != "Chihuahua" , Entidad != "Coahuila" , Entidad != "Nuevo León" , Entidad != "Tamaulipas")
nf = mutate(arrange(nf, desc(Distribucion)), n = 1:nrow(nf), num = num_cardinales[1:nrow(nf)]) #Ranking estados no fronterizos


#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
cj = round(filter(dis_tot_est, Entidad == "Jalisco")[,2],1) #Concentración est Jalisco
lj = filter(re, Entidad == "Jalisco")[,4] #Lugar Jalisco núm cardinal
nj = filter(re, Entidad == "Jalisco")[,3] #Lugar Jalisco número
ljn = filter(nf, Entidad == "Jalisco")[,4] #Lugar Jalisco no fronterizo núm cardinal
njn = filter(nf, Entidad == "Jalisco")[,3] #Lugar Jalisco no fronterizo número


descripcion_5 <- paste0("Durante ", mesp, " de ", año, ", Jalisco concentró ",
                        cj, "% del total de los establecimientos con programa ",
                        "IMMEX (manufactureros y no manufactureros), ubicando a ",
                        "Jalisco en el ", lj, " lugar a nivel nacional, con ",
                        "respecto a las demás entidades federativas participantes en el programa.")

if(njn == 1){
  descripcion_5 <- paste0(descripcion_5, " Jalisco se mantiene como el estado no fronterizo",
                          " con la concentración más alta de estos establecimientos,",
                          " por encima de entidades como ", nf[njn+1,1], ", ", nf[njn+2,1],
                          ", ", nf[njn+3,1], " y ", nf[njn+4,1], ".")
}



# TEXTO DESCRIPCIÓN 6: tot_tra, jal_datacu_var ------------------------

#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
ud = tot_tra[nrow(tot_tra), 3] #Último dato
md = tot_tra[nrow(tot_tra)-1,3] #Dato mes anterior
ad = tot_tra[nrow(tot_tra)-12,3] #Dato año anterior
uv = round(jal_datacu_var[nrow(jal_datacu_var),17],1) #Última variación
av = round(jal_datacu_var[nrow(jal_datacu_var)-1,17],1) #Última variación mes
up = tot_tra[nrow(tot_tra),4] #Último promedio
ap = tot_tra[nrow(tot_tra)-1,4] #Promedio anterior

descripcion_6 <- "El personal ocupado en el total de los establecimientos manufactureros y no manufactureros con programa IMMEX en Jalisco "

if(ud > md){
  descripcion_6 <- paste0(descripcion_6, "aumentó de ", format(md, big.mark = ","), " a ", format(ud, big.mark = ","), " en ", 
                          mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else if(ud < md){
  descripcion_6 <- paste0(descripcion_6, "disminuyó de ", format(md, big.mark = ","), " a ", format(ud, big.mark = ","), " en ", 
                          mesp, " de ", año, " respecto al mes inmediato anterior. ")
}else{
  descripcion_6 <- paste0(descripcion_6, " fue de ", format(ud, big.mark = ","), " trabajadores ",
                          "en línea con la cifra del mes inmediato anterior. ")
}


#conector

if(ud > md){
  if(ud > ad){
    conector <- "Además, "
  }else{
    conector <- "Sin embargo, "
  }
}else{
    if(ud < ad | ud == ad){
      conector <- "Además, "
    }else{
      conector <- "Sin embargo, "
    }
  }

if(ud > ad){
  descripcion_6 <- paste0(descripcion_6, conector, "esta cifra fue superior a la de ",
                          mesp, " de ", año-1, ", la cual se ubicaba en ", format(ad, big.mark = ","),
                          " personas ocupadas. Esto representó un crecimiento anual de ", uv,
                          "%")
}else if(ud < ad){
  descripcion_6 <- paste0(descripcion_6, conector, "esta cifra fue inferior a la de ",
                          mesp, " de ", año-1, ", la cual se ubicaba en ", format(ad, big.mark = ","),
                          " personas ocupadas. Esto representó una reducción anual de ", abs(uv),
                          "%")
}else{
  descripcion_6 <- paste0(descripcion_6, conector, "esta cifra se mantuvo en línea con la de ", mesp,
                          " de ", año-1)
}

##-- 3 escenarios uv > av y 3 escenarios uv < av --##
  if(uv > av){
    if(uv > 0 & av > 0){
      descripcion_6 <- paste0(descripcion_6, ", crecimiento superior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv > 0 & av < 0){
      descripcion_6 <- paste0(descripcion_6, ", cifra superior a la observada el mes anterior, ",
                              "cuando se presentó una disminución de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv < 0 & av < 0){
      descripcion_6 <- paste0(descripcion_6, ", caída menor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }
  }else if(uv < av){
    if(uv > 0 & av > 0){
      descripcion_6 <- paste0(descripcion_6, ", crecimiento inferior al observado el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv < 0 & av > 0){
      descripcion_6 <- paste0(descripcion_6, ", cifra inferior a la observada el mes anterior, ",
                              "cuando se presentó un crecimiento de ", format(abs(av),nsmall = 1), "% anual.")
    }else if(uv < 0 & av < 0){
      descripcion_6 <- paste0(descripcion_6, ", caída mayor a la observada el mes anterior, ",
                              "que fue de ", format(abs(av),nsmall = 1), "% anual.")
    }
  }else{
    descripcion_6 <- paste0(descripcion_6, ", en línea con la cifra observada el mes anterior.")
  }

descripcion_6_2 <- "El promedio de los últimos 12 meses "

#Promedio
if(up > ap){#Aumenta promedio
  ac = 0 #aumento consecutivo
  while(tot_tra[nrow(tot_tra)-ac,4] > (tot_tra[nrow(tot_tra)-1-ac,4])){
    ac = ac + 1
  }
  dcp = 0 #descenso consecutivo previo
  while(tot_tra[nrow(tot_tra)-ac-1-dcp,4] < tot_tra[nrow(tot_tra)-ac-2-dcp,4]){
    dcp = dcp + 1
  }
  dcp = dcp + 1
  
  if(ac == 1 & dcp <3){
    descripcion_6_2 = paste0(descripcion_6_2,"registró un aumento")
  }else if(ac == 1 & dcp >= 3){
    descripcion_6_2 = paste0(descripcion_6_2,"registró su primer aumento después de ",dcp,
                            " meses consecutivos de descensos")
  }else if(ac > 1 & ac <= 20){
    descripcion_6_2 = paste0(descripcion_6_2,"aumentó por ",num_cardinales[ac]," mes consecutivo")
  }else if(ac > 1 & ac > 20){
    descripcion_6_2 = paste0(descripcion_6_2,"mantiene su tendencia creciente")
  }
  descripcion_6_2 = paste0(descripcion_6_2,", al pasar de ",format(ap,big.mark=","), " a ",format(up,big.mark=","), " ocupados en ",mesp, " de ",
                          año, " respecto al mes inmediato anterior.")
}else if( up < ap){#Disminuye promedio
  dc = 0 #descenso consecutivo
  while(tot_tra[nrow(tot_tra)-dc,4] < (tot_tra[nrow(tot_tra)-1-dc,4])){
    dc = dc + 1
  }
  acp = 0 #ascenso consecutivo previo
  while(tot_tra[nrow(tot_tra)-dc-1-acp,4] > tot_tra[nrow(tot_tra)-dc-2-acp,4]){
    acp = acp + 1
  }
  acp = acp + 1
  
  if(dc == 1 & acp <3){
    descripcion_6_2 = paste0(descripcion_6_2,"registró un descenso")
  }else if(dc == 1 & acp >= 3){
    descripcion_6_2 = paste0(descripcion_6_2,"registró su primer descenso después de ",acp,
                            " meses consecutivos de incrementos")
  }else if(dc > 1 & dc <= 20){
    descripcion_6_2 = paste0(descripcion_6_2,"disminuyó por ",num_cardinales[dc]," mes consecutivo")
  }else if(dc > 1 & dc > 20){
    descripcion_6_2 = paste0(descripcion_6_2,"mantiene su tendencia decreciente")
  }
  descripcion_6_2 = paste0(descripcion_6_2,", al pasar de ",format(ap,big.mark=","), " a ",format(up,big.mark=","), " ocupados en ",mesp, " de ",
                          año, " respecto al mes inmediato anterior.")
}else{
  descripcion_6_2 = paste0(descripcion_6_2,conector1,"se mantuvo sin cambios.")
}



# TEXTO DESCRIPCIÓN 7: jal_datacu_var & jal_datacu_var_pro ----------------

#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
uv = round(jal_datacu_var[nrow(jal_datacu_var),17],1) #Última variación
av = round(jal_datacu_var[nrow(jal_datacu_var)-12,17],1) #Variación año anterior
up = round(jal_datacu_var_pro[nrow(jal_datacu_var_pro),17],1) #Última var pro
ap = round(jal_datacu_var_pro[nrow(jal_datacu_var_pro)-1,17],1) #Variación prom mes anterior

##----- Por otra parte, -----##
descripcion_7 <- "Por otra parte, "
if(uv >= 0){
  if(uv > av & av > 0){
    descripcion_7 <- paste0(descripcion_7,"el crecimiento anual del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
    descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_7 <- paste0(descripcion_7, "fue superior al de ", mesp, " de ", año-1, ", ")
    descripcion_7 <- paste0(descripcion_7, "cuando se presentó un incremento de ", format(abs(av),nsmall = 1), "%. ")
    }else if(uv > av & av < 0){
      descripcion_7 <- paste0(descripcion_7,"la variación anual del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
      descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
      descripcion_7 <- paste0(descripcion_7, "fue superior a la de ", mesp, " de ", año-1, ", ")
      descripcion_7 <- paste0(descripcion_7, "cuando se presentó una disminución de ", format(abs(av),nsmall = 1), "%. ")
    }else if(uv > av & av == 0){
      descripcion_7 <- paste0(descripcion_7,"el crecimiento del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
      descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
      descripcion_7 <- paste0(descripcion_7, "fue superior al de ", mesp, " de ", año-1, ", ")
      descripcion_7 <- paste0(descripcion_7, "cuando no se presentó variación respecto a ", mesp, " de ", año-2, ". ")
    }else if(uv < av){
      descripcion_7 <- paste0(descripcion_7,"el crecimiento del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
      descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
      descripcion_7 <- paste0(descripcion_7, "fue inferior al de ", mesp, " de ", año-1, ", ")
      descripcion_7 <- paste0(descripcion_7, "cuando se presentó un incremento de ", format(abs(av),nsmall = 1), "%. ")
    }
}else if(uv < 0){
  if(uv > av){
    descripcion_7 <- paste0(descripcion_7,"la disminución anual del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
    descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_7 <- paste0(descripcion_7, "fue menor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_7 <- paste0(descripcion_7, "cuando se presentó una reducción de ", format(abs(av),nsmall = 1), "%. ")
  }else if(uv < av & av < 0){
    descripcion_7 <- paste0(descripcion_7,"la disminución  anual del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
    descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(abs(uv),nsmall = 1), "%, ")
    descripcion_7 <- paste0(descripcion_7, "fue mayor a la caída de ", mesp, " de ", año-1, ", ")
    descripcion_7 <- paste0(descripcion_7, "cuando se presentó una reducción de ", format(abs(av),nsmall = 1), "%. ")
  }else if(uv < av & av >= 0){
    descripcion_7 <- paste0(descripcion_7,"la variación del personal ocupado en establecimientos en el programa IMMEX en Jalisco de ")
    descripcion_7 <- paste0(descripcion_7, mesp, " de ", año, ", de ", format(uv,nsmall = 1), "%, ")
    descripcion_7 <- paste0(descripcion_7, "fue inferior a la de ", mesp, " de ", año-1, ", ")
    descripcion_7 <- paste0(descripcion_7, "cuando se presentó un crecimiento de ", format(av,nsmall = 1), "%. ")
  }
}


#Conector
if(uv > av){
  if(uv > up){
    conector <- "Además, "
  }else{
    conector <- "Sin embargo, "
  }
}else{
  if(uv <= up){
    conector <- "Además, "
  }else{
    conector <- "Sin embargo, "
  }
}

descripcion_7 <- paste0(descripcion_7, conector, "la última cifra de variación se ubicó ")

if(uv > up){
  descripcion_7 <- paste0(descripcion_7, "por arriba del ")
}else if(uv < up){
  descripcion_7 <- paste0(descripcion_7, "por abajo del ")
}else{
  descripcion_7 <- paste0(descripcion_7, "en línea con el ")
}

descripcion_7 <- paste0(descripcion_7, "promedio de los últimos doce meses, mientras que la variación promedio ")

if(up == ap){
  descripcion_7 <- paste0(descripcion_7, " anual en ", mesp, " de ", año, " se mantuvo constante respecto a la del mes inmediato anterior.")
}else{
  descripcion_7 <- paste0(descripcion_7, "pasó de ", format(ap ,nsmall = 1), "% a ", format(up ,nsmall = 1), "% anual en ",
                          mesp, " de ", año, " respecto al mes inmediato anterior.")
}



# TEXTO DESCRIPCIÓN 8: dis_tot_tra ----------------------------------------

#Dataframes
re = mutate(arrange(dis_tot_tra, desc(Distribucion)), n = 1:nrow(dis_tot_tra), num = num_cardinales[1:nrow(dis_tot_tra)])#Ranking estados
nf = filter(dis_tot_tra, Entidad != "Baja California" , Entidad != "Sonora" , Entidad != "Chihuahua" , Entidad != "Coahuila" , Entidad != "Nuevo León" , Entidad != "Tamaulipas")
nf = mutate(arrange(nf, desc(Distribucion)), n = 1:nrow(nf), num = num_cardinales[1:nrow(nf)]) #Ranking estados no fronterizos

#variables
mesactual = meses[as.numeric(substr(Sys.Date(),6,7)),1]
añoactual = as.numeric(substr(Sys.Date(),1,4))
mesp = meses[mes,1]
cj = round(filter(dis_tot_tra, Entidad == "Jalisco")[,2],1) #Concentración est Jalisco
lj = filter(re, Entidad == "Jalisco")[,4] #Lugar Jalisco núm cardinal
nj = filter(re, Entidad == "Jalisco")[,3] #Lugar Jalisco número
ljn = filter(nf, Entidad == "Jalisco")[,4] #Lugar Jalisco no fronterizo núm cardinal
njn = filter(nf, Entidad == "Jalisco")[,3] #Lugar Jalisco no fronterizo número

descripcion_8 <- paste0("Durante ", mesp, " de ", año, ", Jalisco concentró ",
                        cj, "% del total del personal ocupado en los establecimientos ",
                        "manufactureros y no manufactureros con programa IMMEX, ",
                        "ubicando a Jalisco en el ", lj, " lugar a nivel nacional, ",
                        "con respecto a las entidades federativas con programa. Jalisco ")


if(njn == 1){
  descripcion_8 <- paste0(descripcion_8, "se mantiene como el estado no fronterizo",
                          " con la concentración más alta de personal ocupado en estos establecimientos,",
                          " por arriba de entidades como ", nf[njn+1,1], ", ", nf[njn+2,1],
                          ", ", nf[njn+3,1], " y ", nf[njn+4,1], ".")
}





########## EXPORTAR A EXCEL ##########
#Fuente y titulos
#ft=read.csv("fuente_titulos_figuras.csv",header = FALSE,stringsAsFactors = FALSE)
ft=data.frame(c("Fuente: IIEG, con información de INEGI.",
                "Ingresos provenientes del mercado extranjero que obtuvieron los establecimientos manufactureros del programa IMMEX en Jalisco. Cifras mensuales en millones de pesos,",
                "Ingresos provenientes del mercado extranjero que obtuvieron los establecimientos no manufactureros en el programa IMMEX en Jalisco. Cifras mensuales en millones de pesos,",
                "Ingresos provenientes del mercado extranjero que obtuvieron los establecimientos manufactureros y no manufactureros en el programa IMMEX en Jalisco. Cifras acumuladas anuales en millones de pesos y su variación anual,",
                "Número de establecimientos manufactureros y no manufactureros en el programa IMMEX en Jalisco,",
                "Distribución porcentual de los establecimientos manufactureros y no manufactureros con programa IMMEX por entidad federativa,",
                "Personal ocupado en establecimientos manufactureros y no manufactureros en el programa IMMEX en Jalisco,",
                "Variación porcentual anual de personal ocupado en establecimientos manufactureros y no manufactureros en el programa IMMEX en Jalisco,",
                "Distribución porcentual del personal ocupado de los establecimientos manufactureros y no manufactureros con programa IMMEX por entidad federativa,"))
colnames(ft)= "FT"

#Notas de graficas
#notas=read.csv("notas_graf.csv",header = FALSE,stringsAsFactors = FALSE)
notas=data.frame("Nota: Las cifras se refieren a la suma de los montos de los últimos 12 meses al mes de referencia.")
colnames(notas) = "Notas"

#Nombres de variables
#nombre_var=read.csv("nombre_variables.csv",header = FALSE,stringsAsFactors = FALSE)
nombre_var=data.frame(c("Año","Mes","Exportaciones","Promedio","Variación","Variación promedio","Establecimientos","Entidad","Distribución porcentual","Personal ocupado"))
colnames(nombre_var)="Variable"
names(man_exp_men_pro)=nombre_var[1:4,1]
names(nom_exp_men_pro)=nombre_var[1:4,1]
names(tot_exp_acu)=nombre_var[c(1:3,5:6),1]
names(tot_est)=nombre_var[c(1,2,7,4),1]
names(dis_tot_est)=nombre_var[c(8,9),1]
names(tot_tra)=nombre_var[c(1,2,10,4),1]
names(var_tot_tra)=nombre_var[c(1,2,5,6),1]
names(dis_tot_tra)=nombre_var[c(8,9),1]


wb=createWorkbook("IIEG DIEEF")
addWorksheet(wb, "MAN EXP MEN")
titulo=paste(ft[2,1],periodo1)
writeData(wb, sheet=1, titulo, startCol=1, startRow=1)
writeData(wb, sheet=1, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=1, man_exp_men_pro, startCol=1, startRow=5)
writeData(wb, sheet=1, man_exp_men_pro_tab, startCol=7, startRow=5)

addWorksheet(wb, "NOM EXP MEN")
titulo=paste(ft[3,1],periodo1)
writeData(wb, sheet=2, titulo, startCol=1, startRow=1)
writeData(wb, sheet=2, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=2, nom_exp_men_pro, startCol=1, startRow=5)
writeData(wb, sheet=2, nom_exp_men_pro_tab, startCol=7, startRow=5)

addWorksheet(wb, "TOT EXP ACU")
titulo=paste(ft[4,1],periodo1)
writeData(wb, sheet=3, titulo, startCol=1, startRow=1)
writeData(wb, sheet=3, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=3, notas[1,1], startCol=1, startRow=3)
writeData(wb, sheet=3, tot_exp_acu, startCol=1, startRow=5)
writeData(wb, sheet=3, tot_exp_acu_tab, startCol=8, startRow=5)

addWorksheet(wb, "TOT EST")
titulo=paste(ft[5,1],periodo1)
writeData(wb, sheet=4, titulo, startCol=1, startRow=1)
writeData(wb, sheet=4, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=4, tot_est, startCol=1, startRow=5)
writeData(wb, sheet=4, tot_est_tab, startCol=7, startRow=5)

addWorksheet(wb, "DIST TOT EST")
titulo=paste(ft[6,1],periodo2)
writeData(wb, sheet=5, titulo, startCol=1, startRow=1)
writeData(wb, sheet=5, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=5, dis_tot_est, startCol=1, startRow=5)

addWorksheet(wb, "TOT TRA")
titulo=paste(ft[7,1],periodo1)
writeData(wb, sheet=6, titulo, startCol=1, startRow=1)
writeData(wb, sheet=6, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=6, tot_tra, startCol=1, startRow=5)
writeData(wb, sheet=6, tot_tra_tab, startCol=7, startRow=5)

addWorksheet(wb, "VAR TOT TRA")
titulo=paste(ft[8,1],periodo1)
writeData(wb, sheet=7, titulo, startCol=1, startRow=1)
writeData(wb, sheet=7, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=7, var_tot_tra, startCol=1, startRow=5)
writeData(wb, sheet=7, var_tot_tra_tab, startCol=7, startRow=5)

addWorksheet(wb, "DIST TOT TRA")
titulo=paste(ft[9,1],periodo2)
writeData(wb, sheet=8, titulo, startCol=1, startRow=1)
writeData(wb, sheet=8, ft[1,1], startCol=1, startRow=2)
writeData(wb, sheet=8, dis_tot_tra, startCol=1, startRow=5)

addWorksheet(wb, "Texto")
titulo="I	Industria Manufacturera, Maquiladora y de Servicios de Exportación, IMMEX"
writeData(wb, sheet = 9, "Título ficha:", startCol = 1, startRow = 1)
writeData(wb, sheet = 9, titulo, startCol = 1, startRow = 2)
###
writeData(wb, sheet = 9, "MAN EXP MEN:", startCol = 1, startRow = 4)
writeData(wb, sheet = 9, "Texto1:", startCol = 1, startRow = 5)
writeData(wb, sheet = 9, "Texto1_2:", startCol = 1, startRow = 6)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 7)
writeData(wb, sheet = 9, descripcion1_1, startCol = 2, startRow = 5)
writeData(wb, sheet = 9, descripcion1_2, startCol = 2, startRow = 6)
writeData(wb, sheet = 9, paste(ft[2,1],periodo1), startCol = 2, startRow = 7)
###
writeData(wb, sheet = 9, "NOM EXP MEN:", startCol = 1, startRow = 9)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 10)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 11)
writeData(wb, sheet = 9, descripcion_2, startCol = 2, startRow = 10)
writeData(wb, sheet = 9, paste(ft[3,1],periodo1), startCol = 2, startRow = 11)
###
writeData(wb, sheet = 9, "TOT EXP ACU:", startCol = 1, startRow = 13)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 14)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 15)
writeData(wb, sheet = 9, "Nota:", startCol = 1, startRow = 16)
writeData(wb, sheet = 9, descripcion_3, startCol = 2, startRow = 14)
writeData(wb, sheet = 9, paste(ft[4,1],periodo1), startCol = 2, startRow = 15)
writeData(wb, sheet = 9, "Nota: Las cifras se refieren a la suma de los montos de los últimos 12 meses al mes de referencia.", startCol = 2, startRow = 16)
###
writeData(wb, sheet = 9, "TOT EST:", startCol = 1, startRow = 18)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 19)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 20)
writeData(wb, sheet = 9, "Nota:", startCol = 1, startRow = 21)
writeData(wb, sheet = 9, descripcion_4, startCol = 2, startRow = 19)
writeData(wb, sheet = 9, paste(ft[5,1],periodo1), startCol = 2, startRow = 20)
writeData(wb, sheet = 9, "Nota: El promedio es el de los últimos 12 meses al mes de referencia.", startCol = 2, startRow = 21)
###
writeData(wb, sheet = 9, "DIST TOT EST:", startCol = 1, startRow = 23)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 24)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 25)
writeData(wb, sheet = 9, descripcion_5, startCol = 2, startRow = 24)
writeData(wb, sheet = 9, paste(ft[6,1],periodo2), startCol = 2, startRow = 25)
###
writeData(wb, sheet = 9, "TOT TRA:", startCol = 1, startRow = 27)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 28)
writeData(wb, sheet = 9, "Texto_2:", startCol = 1, startRow = 29)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 30)
writeData(wb, sheet = 9, "Nota:", startCol = 1, startRow = 31)
writeData(wb, sheet = 9, descripcion_6, startCol = 2, startRow = 28)
writeData(wb, sheet = 9, descripcion_6_2, startCol = 2, startRow = 29)
writeData(wb, sheet = 9, paste(ft[7,1],periodo1), startCol = 2, startRow = 30)
writeData(wb, sheet = 9, "Nota: El promedio se refiere al de los último doce meses.", startCol = 2, startRow = 31)
###
writeData(wb, sheet = 9, "VAR TOT TRA:", startCol = 1, startRow = 33)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 34)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 35)
writeData(wb, sheet = 9, descripcion_7, startCol = 2, startRow = 34)
writeData(wb, sheet = 9, paste(ft[8,1],periodo1), startCol = 2, startRow = 35)
###
writeData(wb, sheet = 9, "DIST TOT TRA:", startCol = 1, startRow = 37)
writeData(wb, sheet = 9, "Texto:", startCol = 1, startRow = 38)
writeData(wb, sheet = 9, "Gráfica:", startCol = 1, startRow = 39)
writeData(wb, sheet = 9, descripcion_8, startCol = 2, startRow = 38)
writeData(wb, sheet = 9, paste(ft[9,1],periodo2), startCol = 2, startRow = 39)



nombre_wb=paste0("IMMEX_R-Excel_",fcsv,".xlsx")
saveWorkbook(wb, nombre_wb, overwrite = TRUE)
