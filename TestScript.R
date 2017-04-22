library(ggplot2)
initial.options <- commandArgs(trailingOnly = FALSE)
script.name <- sub("--file=", "", initial.options[grep("--file=", initial.options)])
script.basename <- dirname(script.name)
if (length(script.basename) == 0) {
  # this only works RGui
  script.basename <- getSrcDirectory(function(x) {x})
}
rawdataFilename<-"test_in.txt"
setwd(script.basename)
print(script.basename)

#rawdata <-as.matrix(read.csv2(rawdataFilename,sep="\t",header=FALSE))
#mode(rawdata) <-"numeric"
rawdata <- read.csv2(rawdataFilename,dec=".",sep="\t",header=TRUE)

write.table(rawdata[,1],file="test_out.txt",sep="\t",row.names=FALSE,col.names=FALSE,quote=FALSE)
gplot <- ggplot(data = rawdata, aes(x = in1, y = in2) ) + geom_point()
png(filename="testdiagram.png")
print(gplot)
dev.off()

