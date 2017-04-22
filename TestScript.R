initial.options <- commandArgs(trailingOnly = FALSE)
script.name <- sub("--file=", "", initial.options[grep("--file=", initial.options)])
script.basename <- dirname(script.name)
rawdataFilename<-"test_in.txt"
setwd(script.basename)
print(script.basename)

rawdata <-as.matrix(read.csv2(rawdataFilename,sep="\t",header=FALSE))
mode(rawdata) <-"numeric"
print(rawdata)

write.table(rawdata[1],file="test_out.txt",sep="\t",row.names=FALSE,col.names=FALSE,quote=FALSE)

