##producto cartesiano:
install.packages("gtools")
install.packages("writexl")
install.packages("readxl")
install.packages("rstudioapi")
install.packages("tidyr")


library(tidyr)
library(gtools)
library(writexl)
library(readxl)
library(rstudioapi)

#fijamos el directorio
carpeta=dirname(rstudioapi::getActiveDocumentContext()$path)
setwd(carpeta)

#Transformar información (base de mercado que nos envían externamente)
##PENDIENTE##


#cargamos la base de datos:
base<-readxl::read_excel("Informe Cierre de Mercado.xlsx", sheet = "USO_JUNIO")

#variables iniciadas 
definir_nominal_recib = 1000000

para_uvr_23<-readxl::read_excel("Informe Cierre de Mercado.xlsx", sheet = "Informe Cierre")
valor_costo_recibido=definir_nominal_recib*(as.numeric(para_uvr_23[25,8]))/100
interes22_recibidos = definir_nominal_recib*0
UVR_hoy <- as.numeric(base[nrow(base),3])
UVR_fin <- as.numeric(base[nrow(base),4])
indexacion_recibida <- (definir_nominal_recib/UVR_hoy)*UVR_fin - (definir_nominal_recib)
indexacion_recibida

remove(para_uvr_23)

#eliminar NA's y columnas no utilizadas
base<-na.omit(base)
eliminar <- c("Vencimiento", "Cupon en 2022")
base<- base[, !(names(base) %in% eliminar)]

#columnas a eliminar en matriz de resultados:
delete <- c("V1",
             "precio1" , 
             "cupon1" , 
             "Factor_indexacion_o_anulacion1" , 
             "Es_UVR1", 
             "Paga_cupon1")

delete2 <- c("V1" , "V2",
             "precio1" , "precio2" ,
             "cupon1" , "cupon2" , 
             "Factor_indexacion_o_anulacion1" , 
             "Factor_indexacion_o_anulacion2" , 
             "Es_UVR1", "Es_UVR2" ,
             "Paga_cupon1" , "Paga_cupon2")

delete3 <- c("V1" , "V2", "V3",
            "precio1" , "precio2" , "precio3",
            "cupon1" , "cupon2" , "cupon3",
            "Factor_indexacion_o_anulacion1" , 
            "Factor_indexacion_o_anulacion2" , 
            "Factor_indexacion_o_anulacion3" , 
            "Es_UVR1", "Es_UVR2" , "Es_UVR3",
            "Paga_cupon1" , "Paga_cupon2" , "Paga_cupon3")


delete4 <- c("V1" , "V2", "V3", "V4",
             "precio1" , "precio2" , "precio3", "precio4",
             "cupon1" , "cupon2" , "cupon3", "cupon4",
             "Factor_indexacion_o_anulacion1" , 
             "Factor_indexacion_o_anulacion2" , 
             "Factor_indexacion_o_anulacion3" ,
             "Factor_indexacion_o_anulacion4" , 
             "Es_UVR1", "Es_UVR2" , "Es_UVR3", "Es_UVR4",
             "Paga_cupon1" , "Paga_cupon2" , "Paga_cupon3" , "Paga_cupon4")



##################
##UN TITULO

nominales1 <- definir_nominal_recib*1 #seq(0.94, 1, 0.01)

matriz1  <- as.data.frame(crossing(nominales1, base))
matriz1


matriz1$valor_a_girar <- valor_costo_recibido - (matriz1$nominales1 * matriz1$precio/100) 

matriz1$efecto_cupon <- (matriz1$nominales1 * matriz1$cupon * matriz1$Paga_cupon) - interes22_recibidos

matriz1$costo_fiscal_neto <- matriz1$valor_a_girar + matriz1$efecto_cupon

matriz1$variacion_saldo_de_la_deuda <- matriz1$nominales1 - definir_nominal_recib

matriz1$efecto_indexacion <- ((matriz1$nominales1/UVR_hoy)*matriz1$Factor_indexacion_o_anulacion) - matriz1$nominales1*matriz1$Es_UVR - indexacion_recibida

matriz1$resultado_general <- matriz1$costo_fiscal_neto + matriz1$variacion_saldo_de_la_deuda + matriz1$efecto_indexacion


matriz1 <- matriz1[matriz1$valor_a_girar>0,]

#ordenar por efecto fiscal neto
matriz1 <- matriz1[,!(names(matriz1) %in% delete)]
matriz_ordenada_1 <- matriz1[order(matriz1$resultado_general),]

#exportar

df1 = write_xlsx(matriz_ordenada_1, "1_titulo.xlsx")


#remove(matriz, matriz_ordenada_1)

##################
#DOS TITULOS


#nominales
nominales2 <- definir_nominal_recib*1 #seq(0.94, 1, 0.01)

#participaciones 
num.monedas2=2
posibles.combinatorias=100000
aciertos=0
matriz.participaciones <- matrix(data = NA ,nrow = posibles.combinatorias,ncol = num.monedas2 )

for (i in seq(20, 80, 1)){
  for (j in seq(20, 80, 1)){
    if (i+j==100){
      aciertos=aciertos+1
      matriz.participaciones[aciertos,1]=i/100
      matriz.participaciones[aciertos,2]=j/100
    }
  }
}

matriz.participaciones2 <- as.data.frame(na.omit(matriz.participaciones))
names(matriz.participaciones2)[1] <- "participacion1"
names(matriz.participaciones2)[2] <- "participacion2"


#combinaciones
vector_de_monedas <- 1:nrow(base)
matriz_combinaciones2 <- as.data.frame(combinations(n=nrow(base), r= num.monedas2, v=vector_de_monedas))

matriz_combinaciones2$tipo1 = ""
matriz_combinaciones2$tipo2 = ""

matriz_combinaciones2$precio1 = 0
matriz_combinaciones2$precio2 = 0

matriz_combinaciones2$cupon1 = 0
matriz_combinaciones2$cupon2 = 0

matriz_combinaciones2$Factor_indexacion_o_anulacion1 = 0
matriz_combinaciones2$Factor_indexacion_o_anulacion2 = 0

matriz_combinaciones2$Es_UVR1 = 0
matriz_combinaciones2$Es_UVR2 = 0

for(i in 1:nrow(matriz_combinaciones2)){
  
  matriz_combinaciones2[i,"tipo1"] <- base[matriz_combinaciones2[i,1],"tipo"]
  matriz_combinaciones2[i,"tipo2"] <- base[matriz_combinaciones2[i,2],"tipo"]
  
  matriz_combinaciones2[i,"precio1"] <- base[matriz_combinaciones2[i,1],"precio"]/100
  matriz_combinaciones2[i,"precio2"] <- base[matriz_combinaciones2[i,2],"precio"]/100
  
  matriz_combinaciones2[i, "cupon1"] <- base[matriz_combinaciones2[i,1],"cupon"]
  matriz_combinaciones2[i, "cupon2"] <- base[matriz_combinaciones2[i,2],"cupon"]
  
  matriz_combinaciones2[i, "Factor_indexacion_o_anulacion1"] = base[matriz_combinaciones2[i,1],"Factor_indexacion_o_anulacion"]
  matriz_combinaciones2[i, "Factor_indexacion_o_anulacion2"] = base[matriz_combinaciones2[i,2],"Factor_indexacion_o_anulacion"]
  
  matriz_combinaciones2[i, "Es_UVR1"] = base[matriz_combinaciones2[i,1],"Es_UVR"]
  matriz_combinaciones2[i, "Es_UVR2"] = base[matriz_combinaciones2[i,2],"Es_UVR"]
  
  matriz_combinaciones2[i, "Paga_cupon1"] = base[matriz_combinaciones2[i,1],"Paga_cupon"]
  matriz_combinaciones2[i, "Paga_cupon2"] = base[matriz_combinaciones2[i,2],"Paga_cupon"]
  
}

#matriz_combinaciones2 <- matriz_combinaciones2[,-c(1:2)]


#producto cartesiano
matriz2  <- as.data.frame(crossing(nominales2, matriz.participaciones2, matriz_combinaciones2))
matriz2


matriz2$valor_a_girar <- valor_costo_recibido - (matriz2$nominales2 * matriz2$participacion1 * matriz2$precio1  + 
                                                  matriz2$nominales2 * matriz2$participacion2 * matriz2$precio2 ) 

matriz2$efecto_cupon <- (matriz2$nominales2 * matriz2$participacion1 * matriz2$cupon1 * matriz2$Paga_cupon1 + 
                          matriz2$nominales2 * matriz2$participacion2 * matriz2$cupon2 * matriz2$Paga_cupon2) - interes22_recibidos

matriz2$costo_fiscal_neto <- matriz2$valor_a_girar +matriz2$efecto_cupon

matriz2$variacion_saldo_de_la_deuda <- matriz2$nominales2 - definir_nominal_recib

matriz2$efecto_indexacion <- ((((matriz2$nominales2*matriz2$participacion1*matriz2$Factor_indexacion_o_anulacion1)/UVR_hoy) - (matriz2$nominales2*matriz2$participacion1*matriz2$Es_UVR1)) + 
                              (((matriz2$nominales2*matriz2$participacion2*matriz2$Factor_indexacion_o_anulacion2)/UVR_hoy) - (matriz2$nominales2*matriz2$participacion2*matriz2$Es_UVR2))) -
                                  indexacion_recibida

matriz2$cfn_e_index <- matriz2$costo_fiscal_neto + matriz2$efecto_indexacion

matriz2$resultado_general <- matriz2$costo_fiscal_neto + matriz2$variacion_saldo_de_la_deuda + matriz2$efecto_indexacion


#valor a girar positivo
matriz2<- matriz2[matriz2$valor_a_girar>=0,]
matriz2 <- matriz2[,!(names(matriz2) %in% delete2)]

#ordenar por resultado general
matriz_ordenada2<-matriz2[order(matriz2$resultado_general),]

#exportar
df2 = write_xlsx(matriz_ordenada2, "2_titulos.xlsx")

#remove(matriz,matriz_ordenada2,matriz.participaciones, matriz.participaciones2, matriz_combinaciones2)

##################
##TRES TITULOS 
#nominales
nominales3 <- definir_nominal_recib*1 #seq(0.94, 1, 0.01)
vector_de_monedas <- 1:nrow(base)

  
#para matriz de 3 titulos a entregar
num.monedas3=3
posibles.combinatorias=1000000
aciertos=0
matriz.participaciones3 <- matrix(data = NA ,nrow = posibles.combinatorias,ncol = num.monedas3)


#tres monedas
#si se pide algo espec?fico, revisar l?mites

for (i in seq(10, 90, 1)){
  for (j in seq(10, 90, 1)){
    for(k in seq(10, 90, 1)){
      if (i+j+k==100){
        aciertos=aciertos+1
        matriz.participaciones3[aciertos,1]=i/100
        matriz.participaciones3[aciertos,2]=j/100
        matriz.participaciones3[aciertos,3]=k/100
      }
    }
  }
}

matriz.participaciones3<-as.data.frame(na.omit(matriz.participaciones3))
names(matriz.participaciones3)[1] <- "participacion1"
names(matriz.participaciones3)[2] <- "participacion2"
names(matriz.participaciones3)[3] <- "participacion3"


#generar matriz de combinaciones
matriz_combinaciones3 <- as.data.frame(combinations(n=nrow(base), r= num.monedas3, v=vector_de_monedas))

matriz_combinaciones3$tipo1 = ""
matriz_combinaciones3$tipo2 = ""
matriz_combinaciones3$tipo3 = ""

matriz_combinaciones3$precio1 = 0
matriz_combinaciones3$precio2 = 0
matriz_combinaciones3$precio3 = 0

matriz_combinaciones3$cupon1 = 0
matriz_combinaciones3$cupon2 = 0
matriz_combinaciones3$cupon3 = 0

matriz_combinaciones3$Factor_indexacion_o_anulacion1 = 0
matriz_combinaciones3$Factor_indexacion_o_anulacion2 = 0
matriz_combinaciones3$Factor_indexacion_o_anulacion3 = 0

matriz_combinaciones3$Es_UVR1 = 0
matriz_combinaciones3$Es_UVR2 = 0
matriz_combinaciones3$Es_UVR3 = 0

matriz_combinaciones3$Paga_cupon1 = 0
matriz_combinaciones3$Paga_cupon2 = 0
matriz_combinaciones3$Paga_cupon3 = 0

for(i in 1:nrow(matriz_combinaciones3)){
  
  matriz_combinaciones3[i,"tipo1"] <- base[matriz_combinaciones3[i,1],"tipo"]
  matriz_combinaciones3[i,"tipo2"] <- base[matriz_combinaciones3[i,2],"tipo"]
  matriz_combinaciones3[i,"tipo3"] <- base[matriz_combinaciones3[i,3],"tipo"]
  
  matriz_combinaciones3[i,"precio1"] <- base[matriz_combinaciones3[i,1],"precio"]/100
  matriz_combinaciones3[i,"precio2"] <- base[matriz_combinaciones3[i,2],"precio"]/100
  matriz_combinaciones3[i,"precio3"] <- base[matriz_combinaciones3[i,3],"precio"]/100
  
  matriz_combinaciones3[i, "cupon1"] <- base[matriz_combinaciones3[i,1],"cupon"]
  matriz_combinaciones3[i, "cupon2"] <- base[matriz_combinaciones3[i,2],"cupon"]
  matriz_combinaciones3[i, "cupon3"] <- base[matriz_combinaciones3[i,3],"cupon"]
  
  matriz_combinaciones3[i, "Factor_indexacion_o_anulacion1"] = base[matriz_combinaciones3[i,1],"Factor_indexacion_o_anulacion"]
  matriz_combinaciones3[i, "Factor_indexacion_o_anulacion2"] = base[matriz_combinaciones3[i,2],"Factor_indexacion_o_anulacion"]
  matriz_combinaciones3[i, "Factor_indexacion_o_anulacion3"] = base[matriz_combinaciones3[i,3],"Factor_indexacion_o_anulacion"]
  
  matriz_combinaciones3[i, "Es_UVR1"] = base[matriz_combinaciones3[i,1],"Es_UVR"]
  matriz_combinaciones3[i, "Es_UVR2"] = base[matriz_combinaciones3[i,2],"Es_UVR"]
  matriz_combinaciones3[i, "Es_UVR3"] = base[matriz_combinaciones3[i,3],"Es_UVR"]
  
  matriz_combinaciones3[i, "Paga_cupon1"] = base[matriz_combinaciones3[i,1],"Paga_cupon"]
  matriz_combinaciones3[i, "Paga_cupon2"] = base[matriz_combinaciones3[i,2],"Paga_cupon"]
  matriz_combinaciones3[i, "Paga_cupon3"] = base[matriz_combinaciones3[i,3],"Paga_cupon"]
  
}

#matriz_combinaciones3 <- matriz_combinaciones3[,-c(1:3)]

#producto cartesiano
matriz3  <- as.data.frame(crossing(nominales3, matriz.participaciones3, matriz_combinaciones3))
matriz3


matriz3$valor_a_girar <- valor_costo_recibido - (matriz3$nominales3 * matriz3$participacion1 * matriz3$precio1  + 
                                                  matriz3$nominales3 * matriz3$participacion2 * matriz3$precio2 +
                                                    matriz3$nominales3 * matriz3$participacion3 * matriz3$precio3) 

matriz3$efecto_cupon <- (matriz3$nominales3 * matriz3$participacion1 * matriz3$cupon1 * matriz3$Paga_cupon1) + 
                          (matriz3$nominales3 * matriz3$participacion2 * matriz3$cupon2 * matriz3$Paga_cupon2) +
                            (matriz3$nominales3 * matriz3$participacion3 * matriz3$cupon3 * matriz3$Paga_cupon3) - interes22_recibidos

matriz3$costo_fiscal_neto <- matriz3$valor_a_girar + matriz3$efecto_cupon

matriz3$variacion_saldo_de_la_deuda <- matriz3$nominales3 - definir_nominal_recib

matriz3$efecto_indexacion <- ((((matriz3$nominales3*matriz3$participacion1*matriz3$Factor_indexacion_o_anulacion1)/UVR_hoy) - (matriz3$nominales3*matriz3$participacion1*matriz3$Es_UVR1)) + 
                              (((matriz3$nominales3*matriz3$participacion2*matriz3$Factor_indexacion_o_anulacion2)/UVR_hoy) - (matriz3$nominales3*matriz3$participacion2*matriz3$Es_UVR2)) +
                                (((matriz3$nominales3*matriz3$participacion3*matriz3$Factor_indexacion_o_anulacion3)/UVR_hoy) - (matriz3$nominales3*matriz3$participacion3*matriz3$Es_UVR3))) - 
                                  indexacion_recibida

matriz3$cfn_e_index <- matriz3$costo_fiscal_neto + matriz3$efecto_indexacion


matriz3$resultado_general <- matriz3$costo_fiscal_neto + matriz3$variacion_saldo_de_la_deuda + matriz3$efecto_indexacion

#eliminar valores a girar negativos
matriz3 <- matriz3[matriz3$valor_a_girar>=0,]
matriz3 <- matriz3[,!(names(matriz3) %in% delete3)]

#ordenar por resultado general
matriz_ordenada3<-matriz3[order(matriz3$resultado_general),]

#exportar primeros 2000 resultados ordenados por efecto fiscal neto
df3 = write_xlsx(matriz_ordenada3[1:2000,], "3_titulos.xlsx")


#remove(matriz,matriz_ordenada3,matriz.participaciones3, matriz_combinaciones3)


##################
#cuatro monedas

#matriz de participaciones:
nominales4 <- definir_nominal_recib*1 #seq(0.94, 1, 0.01)
num.monedas4=4
posibles.combinatorias=10000000
aciertos=0
matriz.participaciones4 <- matrix(data = NA ,nrow = posibles.combinatorias,ncol = num.monedas4 )


for (i in seq(10, 90, 5)){
  for (j in seq(10, 90, 5)){
    for(k in seq(10, 90, 5)){
      for(l in seq(10, 90, 5)){
        if (i+j+k+l==100){
          aciertos=aciertos+1
          matriz.participaciones4[aciertos,1]=i/100
          matriz.participaciones4[aciertos,2]=j/100
          matriz.participaciones4[aciertos,3]=k/100
          matriz.participaciones4[aciertos,4]=l/100
        }
      }
    }
  }
}

matriz.participaciones4<-as.data.frame(na.omit(matriz.participaciones4))
names(matriz.participaciones4)[1] <- "participacion1"
names(matriz.participaciones4)[2] <- "participacion2"
names(matriz.participaciones4)[3] <- "participacion3"
names(matriz.participaciones4)[4] <- "participacion4"

#generar matriz de combinaciones
matriz_combinaciones4 <- as.data.frame(combinations(n=nrow(base), r= num.monedas4, v=vector_de_monedas))


matriz_combinaciones4$tipo1 = ""
matriz_combinaciones4$tipo2 = ""
matriz_combinaciones4$tipo3 = ""
matriz_combinaciones4$tipo4 = ""

matriz_combinaciones4$precio1 = 0
matriz_combinaciones4$precio2 = 0
matriz_combinaciones4$precio3 = 0
matriz_combinaciones4$precio4 = 0

matriz_combinaciones4$cupon1 = 0
matriz_combinaciones4$cupon2 = 0
matriz_combinaciones4$cupon3 = 0
matriz_combinaciones4$cupon4 = 0

matriz_combinaciones4$Factor_indexacion_o_anulacion1 = 0
matriz_combinaciones4$Factor_indexacion_o_anulacion2 = 0
matriz_combinaciones4$Factor_indexacion_o_anulacion3 = 0
matriz_combinaciones4$Factor_indexacion_o_anulacion4 = 0

matriz_combinaciones4$Es_UVR1 = 0
matriz_combinaciones4$Es_UVR2 = 0
matriz_combinaciones4$Es_UVR3 = 0
matriz_combinaciones4$Es_UVR4 = 0

matriz_combinaciones4$Paga_cupon1 = 0
matriz_combinaciones4$Paga_cupon2 = 0
matriz_combinaciones4$Paga_cupon3 = 0
matriz_combinaciones4$Paga_cupon4 = 0

for(i in 1:nrow(matriz_combinaciones4)){
  
  matriz_combinaciones4[i,"tipo1"] <- base[matriz_combinaciones4[i,1],"tipo"]
  matriz_combinaciones4[i,"tipo2"] <- base[matriz_combinaciones4[i,2],"tipo"]
  matriz_combinaciones4[i,"tipo3"] <- base[matriz_combinaciones4[i,3],"tipo"]
  matriz_combinaciones4[i,"tipo4"] <- base[matriz_combinaciones4[i,4],"tipo"]
  
  matriz_combinaciones4[i,"precio1"] <- base[matriz_combinaciones4[i,1],"precio"]/100
  matriz_combinaciones4[i,"precio2"] <- base[matriz_combinaciones4[i,2],"precio"]/100
  matriz_combinaciones4[i,"precio3"] <- base[matriz_combinaciones4[i,3],"precio"]/100
  matriz_combinaciones4[i,"precio4"] <- base[matriz_combinaciones4[i,4],"precio"]/100
  
  matriz_combinaciones4[i, "cupon1"] <- base[matriz_combinaciones4[i,1],"cupon"]
  matriz_combinaciones4[i, "cupon2"] <- base[matriz_combinaciones4[i,2],"cupon"]
  matriz_combinaciones4[i, "cupon3"] <- base[matriz_combinaciones4[i,3],"cupon"]
  matriz_combinaciones4[i, "cupon4"] <- base[matriz_combinaciones4[i,4],"cupon"]
  
  matriz_combinaciones4[i, "Factor_indexacion_o_anulacion1"] = base[matriz_combinaciones4[i,1],"Factor_indexacion_o_anulacion"]
  matriz_combinaciones4[i, "Factor_indexacion_o_anulacion2"] = base[matriz_combinaciones4[i,2],"Factor_indexacion_o_anulacion"]
  matriz_combinaciones4[i, "Factor_indexacion_o_anulacion3"] = base[matriz_combinaciones4[i,3],"Factor_indexacion_o_anulacion"]
  matriz_combinaciones4[i, "Factor_indexacion_o_anulacion4"] = base[matriz_combinaciones4[i,4],"Factor_indexacion_o_anulacion"]
  
  matriz_combinaciones4[i, "Es_UVR1"] = base[matriz_combinaciones4[i,1],"Es_UVR"]
  matriz_combinaciones4[i, "Es_UVR2"] = base[matriz_combinaciones4[i,2],"Es_UVR"]
  matriz_combinaciones4[i, "Es_UVR3"] = base[matriz_combinaciones4[i,3],"Es_UVR"]
  matriz_combinaciones4[i, "Es_UVR4"] = base[matriz_combinaciones4[i,4],"Es_UVR"]
  
  matriz_combinaciones4[i, "Paga_cupon1"] = base[matriz_combinaciones4[i,1],"Paga_cupon"]
  matriz_combinaciones4[i, "Paga_cupon2"] = base[matriz_combinaciones4[i,2],"Paga_cupon"]
  matriz_combinaciones4[i, "Paga_cupon3"] = base[matriz_combinaciones4[i,3],"Paga_cupon"]
  matriz_combinaciones4[i, "Paga_cupon4"] = base[matriz_combinaciones4[i,4],"Paga_cupon"]
  
}

#matriz_combinaciones4 <- matriz_combinaciones4[,-c(1:4)]

#producto cartesiano

matriz4  <- as.data.frame(crossing(nominales4, matriz.participaciones4, matriz_combinaciones4))
matriz4


matriz4$valor_a_girar <- valor_costo_recibido - (matriz4$nominales4 * matriz4$participacion1 * matriz4$precio1  + 
                                                  matriz4$nominales4 * matriz4$participacion2 * matriz4$precio2 +
                                                  matriz4$nominales4 * matriz4$participacion3 * matriz4$precio3 +
                                                  matriz4$nominales4 * matriz4$participacion4 * matriz4$precio4) 

matriz4$efecto_cupon <- (matriz4$nominales4 * matriz4$participacion1 * matriz4$cupon1 * matriz4$Paga_cupon1 + 
                          matriz4$nominales4 * matriz4$participacion2 * matriz4$cupon2 * matriz4$Paga_cupon2 +
                          matriz4$nominales4 * matriz4$participacion3 * matriz4$cupon3 * matriz4$Paga_cupon3 + 
                          matriz4$nominales4 * matriz4$participacion4 * matriz4$cupon4 * matriz4$Paga_cupon4) - interes22_recibidos

matriz4$costo_fiscal_neto <- matriz4$valor_a_girar +matriz4$efecto_cupon

matriz4$variacion_saldo_de_la_deuda <- matriz4$nominales4 - definir_nominal_recib

matriz4$efecto_indexacion <- matriz4$nominales4*matriz4$participacion1*(matriz4$Factor_indexacion_o_anulacion1/UVR_hoy - matriz4$Es_UVR1) + 
                              matriz4$nominales4*matriz4$participacion2*(matriz4$Factor_indexacion_o_anulacion2/UVR_hoy - matriz4$Es_UVR2) +
                                matriz4$nominales4*matriz4$participacion3*(matriz4$Factor_indexacion_o_anulacion3/UVR_hoy - matriz4$Es_UVR3) +
                                  matriz4$nominales4*matriz4$participacion4*(matriz4$Factor_indexacion_o_anulacion4/UVR_hoy - matriz4$Es_UVR4) - 
                                    indexacion_recibida

matriz4$cfn_e_index <- matriz4$costo_fiscal_neto + matriz4$efecto_indexacion

matriz4$resultado_general <- matriz4$costo_fiscal_neto + matriz4$variacion_saldo_de_la_deuda + matriz4$efecto_indexacion

#quitar columnas innecesarias
matriz4 <- matriz4[,!(names(matriz4) %in% delete4)]

matriz_ordenada4 <- matriz4[order(matriz4$resultado_general),]

#tomar las primeras 2000 filas y exportar
df4 = write_xlsx(matriz_ordenada4[1:2000,], "4_titulos.xlsx")
#remove(matriz,matriz_ordenada4,matriz.participaciones4, matriz_combinaciones4)




#EXPORTAR RESULTADO EN UN UNICO EXCEL EN DIFERENTES HOJAS

install.packages("openxlsx")
library(openxlsx)
Sys.setenv(JAVA_HOME="")
install.packages("rJava")
library(rJava)
install.packages("xlsx")
library(xlsx)


#Solucionador al problema de dataframe
options(java.parameters = "-Xmx8000m")
jgc <- function()
{
  .jcall("java/lang/System", method = "gc")
}  

wb <- createWorkbook()

for(i in seq_along(matriz_ordenada2))
{
  gc()
  jgc()
  message("Creating sheet", i)
  sheet <- createSheet(wb, sheetName = names(matriz_ordenada2)[i])
  message("Adding data frame", i)
  addDataFrame(matriz_ordenada2[[i]], sheet)
}

#Transformando

Destino <- dirname(rstudioapi::getActiveDocumentContext()$path)
write.xlsx2(matriz_ordenada_1, paste0(Destino, "CANJESFINAL.xlsx"), row.names = F, sheetName = "UN_TITULO")
write.xlsx2(matriz_ordenada2, paste0(Destino, "CANJESFINAL.xlsx"), row.names = F, sheetName = "DOS_TITULOS", append = T)
write.xlsx2(matriz_ordenada3[1:2000,], paste0(Destino, "CANJESFINAL.xlsx"), row.names = F, sheetName = "TRES_TITULOS", append = T)
write.xlsx2(matriz_ordenada4[1:2000,], paste0(Destino, "CANJESFINAL.xlsx"), row.names = F, sheetName = "CUATRO_TITULOS", append = T)


##SHINY###



#================#
#### Paquetes ####
#================#

library(shiny)       # Para poder hacer la aplicación
library(shinythemes) # Para acceder a temas predeterminados
library(readxl)      # Lectura de archivos de Excel
library(tidyverse)   # Manejo de datos y visualización
library(plotly)      # Visualización interactiva
library(agricolae)   # Tablas de frecuencias agrupadas


#=================#
#### Preámbulo ####
#=================#

# Cargar los datos DE CANJES CONSOLIDADOS
canjes = read_excel(file.choose())

#visualizar la base de datos
view(canjes)

# Renombrar variables para su visualización
canjes = canjes %>%
  rename(Nominales = nominales1,
         Tipo_Canje = tipo,
         Cupón = cupon,
         Paga = Paga_cupon,
         Valor_nominal = `Valor Nominal*`,
         Precio = precio,
         Factor = Factor_indexacion_o_anulacion,
         UVR = Es_UVR,
         Valor_de_giro = valor_a_girar,
         Efecto_cupón = efecto_cupon,
         Costo_fiscal_neto = costo_fiscal_neto,
         Variación_saldo_deuda = variacion_saldo_de_la_deuda,
         Efecto_indexación = efecto_indexacion,
         Resultado_general = resultado_general)
# Variables numéricas
canjes_num = canjes %>%
  select(where(is.numeric))

# Regla de Sturges
interv_sturges = sturges.freq(canjes$Nominales)$classes


#=================#
#### Shiny App ####
#=================#


ui = fluidPage(
  
  # Definir tema
  theme = shinytheme('superhero'),
  
  # Encabezado
  fluidRow(column(1, tags$img(src = 'LogoMHCP.png',
                              width = '100px', 
                              height = '100px')),
           column(8, h1(strong('Shiny for Data Analytics - CANJES DE DEUDA')))),
  
  
  # Título de la página
  headerPanel('Canjes de Deuda'),
  
  tabsetPanel(
    type = 'pills',
    
    # Primera pestaña
    tabPanel(
      title = 'Distribución canjes',
      position = 'right',
      # Tipo de estructura de la página
      sidebarLayout(
        
        # Panel 1
        wellPanel(
          
          # Widget desplegable para secleccionar la variable
          selectInput(inputId = 'variable',
                      label = 'Seleccione una variable',
                      choices = colnames(canjes_num),
                      selected = 'Total'),
          
          # Widget deslizante para seleccionar el número de intervalos
          sliderInput(inputId = 'intervalo',
                      label = 'Número de intervalos',
                      min = 94,
                      max = 100,
                      value = interv_sturges
          ),
          
          # Widget checkeable para determinar si usar densidad o no
          checkboxInput(inputId = 'density',
                        label = '¿Kernel de densidad?',
                        value = F),
          
          # Widget para oprimir el botón de acción
          actionButton('boton_aceptar', 
                       'Aceptar',
                       class = 'btn-success')
        ),
        
        
        # Panel 2
        splitLayout(
          
          # Gráfico
          plotlyOutput(outputId = 'histograma'),
          # Tabla
          tableOutput(outputId = 'tabla'),
          
          
        ) # Cierra el segundo panel
        
      ) # Cierra el sidebarLayout
      
    ), # Cierra la primera pestaña
    
    
    tabPanel(
      title = 'Frecuencia',
      
      numericInput(inputId = 'n',
                   'Tamaño de muestra',
                   value = 25),
      
      plotOutput(outputId = 'hist_ejemplo')
      
      
    ) # Cierra segunda pestaña
    
  ) # Cierra todas las pestañas
  
) # Cierra el ui


server = function(input, output){
  
  # Generar un dataframe con la variable seleccionada como objeto reactivo
  variable_num = reactive({
    canjes_num %>% 
      select(input$variable) %>% 
      rename(Var_select = input$variable)
  })
  
  # Generar un vector con la variable seleccionada como objeto reactivo
  var_hist = reactive({
    
    canjes_num[[input$variable]]
    
  })
  
  # Generar la opción de intervalo como objeto reactivo
  intervalo_react = reactive({
    
    input$intervalo
    
    
  })
  
  
  # Crear histograma
  output$histograma = renderPlotly({
    
    # Activar el botón de acción
    input$boton_aceptar
    
    # Condición si el widget checkeable está activado
    if (isolate(input$density) == T){ # isolate() para que la casilla se active con el botón
      grafica = isolate(ggplot(variable_num()) + # isolate() para que la gráfica se active con el botón
                          aes(x = Var_select) + 
                          geom_histogram(aes(y = ..density..,
                                             fill=..count..),
                                         col = 'black',
                                         # Los breaks dependen de los intervalos seleccionados por el usuario
                                         bins = input$intervalo) + 
                          # Graficar densidad
                          geom_density(colour = 'red',
                                       lwd = 0.75) +
                          # El título depende de la variable seleccionada por el usuario
                          labs(title = paste('Distribución de', input$variable),
                               x = 'Puntaje Global',
                               y = '') +
                          theme(legend.position = 'none'))
      # Condición si el widget checkeable NO está activado
    } else {
      grafica = isolate(ggplot(variable_num()) + # isolate() para que la gráfica se active con el botón
                          aes(x = Var_select) + 
                          geom_histogram(aes(y = ..density..,
                                             fill=..count..),
                                         col = 'black',
                                         bins = input$intervalo) +
                          # NOTE QUE NO SE GRAFICA LA DENSIDAD  
                          labs(title = paste('Distribución de', input$variable),
                               x = 'Global',
                               y = '') +
                          theme(legend.position = 'none'))
      
      
    }
    
    # El gráfico que se muestra en la página es un ggplotly
    ggplotly(grafica)
    
  })
  
  
  # Crear la tabla de frecuencia de datos agrupados
  output$tabla = renderTable({
    
    # Activar el botón de acción
    input$boton_aceptar
    
    # Crear el histograma base para la tabla
    # Note que se usa isolate() para que funcione con el botón
    histo = isolate(hist(var_hist(), breaks = intervalo_react()))
    # Tabla de frecuencia
    table.freq(histo)
    
  })
  
  
  
  output$hist_ejemplo = renderPlot({
    
    hist(rnorm(input$n))
    
  })
  
  
}

shinyApp(ui, server)




