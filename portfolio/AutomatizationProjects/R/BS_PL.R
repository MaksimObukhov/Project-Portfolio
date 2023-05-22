library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)


input_name <- list.files(path = 'bs_overview/data/input/', pattern = '*.xlsx')
input_bs <- as.data.frame(read_excel(paste0('bs_overview/data/input/', input_name[1]), col_names = F))
input_fc <- as.data.frame(read_excel(paste0('bs_overview/data/input/', input_name[2]), col_names = F))

exmpl <- read_excel('bs_overview/data/example_excel.xlsx', sheet = 2, col_names = F)
colnames <- read.csv('bs_overview/data/colnames.csv')[,2]

map_ud2 <- read_excel('bs_overview/data/Map_Accounts_KPIs_ReCo_to_OS_20220413.xlsx')[-1,c(3,5)]
ud2_shorts <- map_ud2[map_ud2[2] == 'mt_short',][1]

get_date <- function() {
  temp <- substr(input_bs[3,1], 9, nchar(input_bs[3,1]))
  temp <- strsplit(temp, ' ')
  months <- c('Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
              'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec')
  mon_num <- match(temp[[1]][1], months)
  year <- temp[[1]][2]
  return(paste0(year, 'M',  mon_num))
}

date <- get_date()

create_row <- function(l, i) {
  entity <- strsplit(l[6,1], ' ')[[1]][2]
  acc <- l[i,1]

  df <- data.frame(
    T.C = date,
    E = entity,
    S.C = 'Act_Mng',
    AMT.ZS = l[i,2],
    A = acc,
    F = get_flow(acc),
    IC = 'None',
    UD1 = 'mng_Reported',
    UD2 = get_ud2(acc),
    UD3 = 'None',
    UD4 = 'None',
    UD5 = 'None',
    UD6 = 'None',
    UD7 = 'None',
    UD8 = 'None',
    V = 'YTD',
    C = 'Local',
    O = 'Import',
    SI = NA)
}


acc_first_letter <- function(acc) {
  return(substr(acc, 1, 1))
}


get_flow <- function(acc) {
  code <- acc_first_letter(acc)
  if (code == 'R' | code == 'C') {
    return ('Y100')
  } else if (code == 'S') {
    return ('None')
  } else {
    return ('F999')
  }
}


get_ud2 <- function(acc) {
  if (acc %in% ud2_shorts[[1]]) {
    return ('mt_short')
  } else {
    return ('None')
  }
}


get_report <- function(l) {
  df_final <- data.frame()
  for (i in 9:nrow(l)) {
    if (l[i,1] != 'EK01000003') {
      df_final <- rbind(df_final, create_row(l, i))
    }
  }

  colnames(df_final) <- colnames
  return(df_final[!is.na(df_final[4]),])
}

bs <- get_report(input_bs)
fc <- get_report(input_fc)

res  <- rbind(bs, fc)

write.xlsx(res, 'report.xlsx')

