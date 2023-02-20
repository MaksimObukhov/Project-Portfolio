library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)
library(readxlsb)


file.list <- list.files(path = "FCFF Visualization/data/", pattern='*.xls*') %>%
  str_sort(numeric = T)


sheets <- lapply(file.list, function(file_name) {
    as.data.frame(
      read_xlsb(paste0("FCFF Visualization/data/", file_name),
                            sheet = "EBITDA",
                            col_names = F))
})

names(sheets) <- file.list


mr_f <- lapply(sheets, function(df) {
    df[4:39,c(2,3,19)] %>%
      mutate_all(~replace(., . == '', NA)) %>%
      drop_na(3)
})


plan <- sheets[[1]][4:39,c(2,3,18)] %>%
  mutate_all(~replace(., . == '', NA)) %>%
  drop_na(3)


plan_by_entity <- plan %>%
  select(2, 3) %>%
  filter(!is.na(plan[2])) %>%
  slice(-c(1, 3, 4))


plan_supply <- plan_by_entity %>%
  filter(plan_by_entity[,1] %in% c('SSE (ex. SSED)', 'EPET', 'DE'),)
plan_supply <- data.frame(ind = 'Supply', value = sum(as.double(plan_supply[,2])))


plan_by_industry <- plan %>%
  select(1, 3) %>%
  filter(!is.na(plan[1])) %>%
  slice(-c(1,2))

plan_by_industry[nrow(plan_by_industry) + 1,] <- plan_supply


fct_by_entities <- lapply(mr_f, function(df) {
  df <- df %>%
    select(2, 3) %>%
    filter(!is.na(df[2]))
  return(df[-1,])
})


supply <- lapply(fct_by_entities, function(df) {
  df <- df %>%
    filter(df[,1] %in% c('SSE (ex. SSED)', 'EPET', 'DE'),)
  df <- data.frame(ind = 'Supply', value = sum(as.double(df[,2])))
})


fct_by_industry <- lapply(mr_f, function(df) {
  df <- df %>%
    select(1, 3) %>%
    filter(!is.na(df[1]))
  return(df[-c(1,2),])
})

for (i in 1:6) {
  fct_by_industry[[i]][nrow(fct_by_industry[[i]]) + 1,] <- supply[[i]]
}


ind_by_month <- plan_by_industry
for (i in 1:6) {
  ind_by_month <- full_join(ind_by_month, fct_by_industry[[i]], by = 'column.2')
}
colnames(ind_by_month) <- c('Industry', '22P', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul')


ent_by_month <- plan_by_entity
for (i in 1:6) {
  ent_by_month <- full_join(ent_by_month, fct_by_entities[[i]], by = 'column.3')
}
colnames(ent_by_month) <- c('Entity', '22P','Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul')


write.xlsx(ent_by_month, 'ent_by_month.xlsx')
write.xlsx(ind_by_month, 'ind_by_month.xlsx')


