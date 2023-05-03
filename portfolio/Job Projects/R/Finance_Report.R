# setwd("C:/Users/EP00345/OneDrive - EPIA/Desktop/R projects/report")

library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)

czk_eur <- 0.040570206


ent_curr <- read_excel('report/data/entity_currency.xlsx')

acc_avoid <- c('SDD0000001', 'AF08000001')

eur_ent <- c()
for (i in seq(nrow(ent_curr))) {
  if (ent_curr[i, 2] == 'EUR') {
    eur_ent <- as.character(append(eur_ent, ent_curr[i, 1]))
  }
}

file.list <- list.files(path = "report/data/", pattern = '*.xlsx')
file.list <- file.list[!grepl('^~', file.list)]

input <- list.files(path = 'report/data/input/H01', pattern = '*.xlsx')

acc_pairs <- data.frame(read_excel('report/data/account_pairs.xlsx'))

keep_codes <- function(xls_name='FRC_partner_IC.xlsx') {
  xls <- as.data.frame(read_excel(paste0('report/data/input/H01/', xls_name), col_names = F))
  code <- NA
  for (i in seq(xls[,1])) {
    if (!is.na(xls[i,1]) & is.na(code)) {
      code <- xls[i,1]
    } else if (is.na(xls[i,1]) & !is.na(code)) {
      xls[i,1] <- code
    } else if (!is.na(xls[i,1]) & !is.na(code)) {
      code <- xls[i,1]
    }
  }
  return(xls)
}


create_row <- function(l, i) {
  entity <- l[1,3]
  acc <- substr(l[i,2], 1, 10)

  df <- data.frame(
    E = entity,
    S.C = l[2,3],
    A = get_counter_acc(acc),
    F = get_flow(acc),  #?
    IC = substr(l[i,1], 1, 4),  #?
    UD1 = 'mng_Reported',
    UD2 = 'None',
    UD3 = 'IC_local',
    UD4 = 'None',
    UD5 = 'None',
    UD6 = 'None',
    UD7 = 'None',
    UD8 = 'None',
    TV_Ann = NA,
    V = 'YTD')

  df <- df %>%
    cbind(get_row(l, i, entity, acc_first_letter(acc), czk_eur)) %>%
    cbind(data.frame(NA, 'IN61', NA, 'Import', NA))

  return(df)
}

# View(create_row(l, 5))

get_counter_acc <- function(acc) {
  l <- acc_pairs
  for (i in seq(l[,1])) {
    if (!is.na(l[i,1])) {
      if (acc == l[i,1]) {
        return (ifelse(!is.na(l[i,2]),
          l[i,2],
          paste('NOPAIRFOUND', acc)))
      }
    }
  }

  for (i in seq(l[,2])) {
    if (!is.na(l[i,2])) {
      if (acc == l[i,2]) {
        return (ifelse(!is.na(l[i,1]),
          l[i,1],
          paste('NOPAIRFOUND', acc)))
      }
    }
  }

  return (paste('NOPAIRFOUND', acc))
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


get_row <- function(l, i, entity, acc_fl, ex_rate) {
  # year <- as.numeric(substr(date, 1, 4))
  # month <- as.numeric(substr(date, 6, 6))

  options(digits=15)
  report_row <- l[i, 3:ncol(l)] %>%
    mutate_all(~replace(., . =='0', NA)) %>%
    mutate_all(~as.numeric(.))

  if (acc_fl %in% c('R', 'C', 'S')) {
    report_row <- as.data.frame(lapply(report_row, function(x) -1 * x))
  }

  if (entity %in% eur_ent) {
    report_row <- as.data.frame(lapply(report_row, function(x) 1/ex_rate * x))
  }

  # report_row <- cbind(data.frame(matrix(NA, nrow = 1, ncol = month - 1)), report_row)
  return(report_row)
}


get_report <- function(l) {
  df_final <- data.frame()
  for (i in 4:length(l[,1])) {
    if (!substr(l[i,2], 1, 10) %in% acc_avoid) {
      df_final <- rbind(df_final, create_row(l, i))
    }
  }
  colnames(df_final) <- as.character(names(read_excel('report/data/FC_IC_Excel.xlsx')))
  return(df_final)
}

l <- keep_codes(input)
rep <- get_report(keep_codes(input))


# write.xlsx(rep, 'report/report.xlsx')
write.csv(rep, 'report/report.csv')


