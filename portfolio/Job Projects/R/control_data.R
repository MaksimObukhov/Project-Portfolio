library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)

# setwd("C:/Users/EP00345/OneDrive - EPIA/Desktop/R projects/control")

rep.list <- list.files(path = "data/reports/", pattern='*Export.xlsx')
other.list <- list.files(path = 'data/', pattern = '*.xlsx')

reports <- lapply(rep.list, function(file_name) {
  as.data.frame(read_excel(paste0("data/reports/", file_name), col_names = F))
})

reports <- lapply(reports, function(rep) {
  rep[, c(1,14)]
})


others <- lapply(other.list, function(file_name) {
  as.data.frame(read_excel(paste0("data/", file_name), col_names = F))
})

others[[1]] <- read_excel(paste0("data/", other.list[1]), sheet = 2, col_names = F)

names(reports) <- rep.list
names(others) <- other.list

fx <- c('FCFF01040T',
        'FCFF0200TT',
        'FCFF0300TT',
        'FCFF0400TT',
        'FCFF0600TT',
        'FCFF0801TT',
        'FCFF020TTT',
        'FCFF01050T',
        'FCFF090201',
        'FCFF090203')


reco <- others[['Reco_OS_EPIF FCFF 2022-03.xlsx']]
reco <- reco[,-1]
rec <- reco[reco[[1]] %in% fx, ]
rec <- rbind(reco[2,], rec)

tr <- t(rec)
tr <- tr[c(-1, -3, -4),]
colnames(tr) <- tr[1,]
colnames(tr)[1] <- 'code'
# tr <- tr[-1,]
tr <- tr[order(tr[,1]),]
ReCo_reports <- as.data.frame(tr)


temp <- as.list(1:length(reports))
for (i in seq(reports)) {
  temp[[i]] <- inner_join(others[[2]][-1,], reports[[i]], by = c('...2' = '...1'))
  temp[[i]] <- temp[[i]][order(temp[[i]][,1]),]
  temp[[i]] <- temp[[i]][-2]
  colnames(temp[[i]]) <- c('code', substr(rep.list[i], 1, 6))
}

names(temp) <- rep.list

options(scipen = 100, digits=10)
# for (i in seq(temp)) {
#   temp[[i]][2] <- temp[[i]][2] %>% mutate(replace(., is.na(.), 0))
#   temp[[i]][,3] <- temp[[i]][,3] %>% as.numeric()
# }

OS_reports <- temp[[1]]
for (i in 2:length(temp)) {
  OS_reports <- inner_join(OS_reports, temp[[i]], by = 'code')
}


ReCo_m <- as.matrix(ReCo_reports[,-1])
OS_m <- as.matrix(OS_reports[,-1])

rc <- matrix(as.numeric(ReCo_m),
            ncol = ncol(ReCo_m))

os <- matrix(as.numeric(OS_m),
             ncol = ncol(OS_m))

div_m <- os - rc
div_m <- as.data.frame(cbind(OS_reports[,1], div_m))
div_m[div_m == 0] <- NA


colnames(div_m) <- colnames(OS_reports)

write.xlsx(ReCo_reports, 'reco_reports.xlsx')

