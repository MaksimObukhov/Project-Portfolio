library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)


input_names <- list.files(path = "joining dat/data/input/", pattern = '*.xlsx')
input <- lapply(input_names, function(input_name) {
    as.data.frame(read_excel(paste0("joining dat/data/input/", input_name), col_names = F))
})

names(input) <- input_names


bs <- full_join(input[[1]], input[[3]], by = 'code')
bs <- full_join(bs, input[[6]], by = 'code')
names(bs) <- c('code', '6370', '7108', '7108GR')

pl <- full_join(input[[2]], input[[4]], by = 'code')
pl <- full_join(pl, input[[5]], by = 'code')
names(pl) <- c('code', '6370', '7108', '7108GR')

bs[,2:4] <- bs[,2:4] %>%
  mutate_all(~replace(., . == '0', NA)) %>%
  mutate_all(~as.numeric(.))

pl[,2:4] <- pl[,2:4] %>%
  mutate_all(~replace(., . == '0', NA)) %>%
  mutate_all(~as.numeric(.))


write.xlsx(pl, 'pl.xlsx')

