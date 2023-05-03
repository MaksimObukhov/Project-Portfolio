library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)

wd <- paste0(getwd(), '/R projects/report/')

file.list <- list.files(path = paste0(wd, 'data/'), pattern = '*.xlsx')
file.list <- file.list[!grepl('^~', file.list)]

input_name <- list.files(path = paste0(wd, 'data/input/O16/'), pattern = '*.xlsx')
input <- as.data.frame(read_excel(paste0(wd, 'data/input/O16/', input_name[1]), col_names = F))

acc_pairs <- data.frame(read_excel(paste0(wd, 'data/account_pairs.xlsx')))

map_ud2 <- read_excel('R projects/bs_overview/data/Map_Accounts_KPIs_ReCo_to_OS_20220413.xlsx')[-1,c(3,5)]
ud2_shorts <- map_ud2[map_ud2[2] == 'mt_short',][1]


date <- '2022M8'
year <- as.numeric(substr(date, 1, 4))
month <- as.numeric(substr(date, 6, 6))
to_CZK <- TRUE
czk_eur <- 0.041025641



ent_curr <- read_excel(paste0(wd,'data/entity_currency.xlsx'))
acc_avoid <- c('SDD0000001', 'AF08000001')

eur_ent <- c()
for (i in seq(nrow(ent_curr))) {
  if (ent_curr[i, 2] == 'EUR') {
    eur_ent <- as.character(append(eur_ent, ent_curr[i, 1]))
  }
}


# l <- input  ##################

create_row <- function(l, i) {
  entity <- strsplit(l[6,1], ' ')[[1]][2]
  acc <- substr(l[i,2], 1, 10)

  df <- data.frame(
    E = entity,
    S.C = get_sc(l[4,1]),
    A = get_counter_acc(acc),
    F = get_flow(acc),
    IC = substr(l[i,1], 1, 4),
    UD1 = 'mng_Reported',
    UD2 = get_ud2(acc), # F999 mt short
    UD3 = 'IC_local',
    UD4 = 'None',
    UD5 = 'None',
    UD6 = 'None',
    UD7 = 'None',
    UD8 = 'None',
    TV_Ann = NA,
    V = 'YTD')

  row <- get_row(l, i, entity, acc_first_letter(acc), to_CZK, czk_eur)
  if (!is.null(row)) {
    df <- cbind(df, row) %>%
    cbind(data.frame(NA, 'IN61', NA, 'Import', NA))
    return(df)
  }

}

get_sc <- function(sc) {
  s <- strsplit(sc, ' ')[[1]]
  paste('Fct',
         substr(s[3], 1, 3),
         s[4],
         substr(s[6], 3, 4),
         sep = '_')
}


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


get_ud2 <- function(acc) {
  if (acc %in% ud2_shorts[[1]]) {
    return ('mt_short')
  } else {
    return ('None')
  }
}

get_row <- function(l, i, entity, acc_fl, to_CZK, ex_rate) {
  options(digits=15)
  report_row <- l[i, 3:ncol(l)]

  colnames(report_row) <- paste(as.character(l[8, 3:ncol(l)]),
                                as.character(l[9, 3:ncol(l)]),
                                sep = '.')

  colnames(report_row) <- lapply(colnames(report_row), function(x) {
    gsub(' ', '.', x)
  })

  report_row <- report_row %>%
    select(contains('Partner')) %>%
    mutate_all(~replace(., . == '0', NA)) %>%
    mutate_all(~as.numeric(.))

  if (all(is.na(report_row))) {
    return (NULL)
  }

  if (acc_fl %in% c('R', 'C', 'S')) {
    report_row <- as.data.frame(lapply(report_row, function(x) -1 * x))
  }

  # report_row <- as.data.frame(lapply(report_row, function(x) 1000 * x)) # kEUR to EUR

  if (to_CZK) {
    if (entity %in% eur_ent){
      report_row <- as.data.frame(lapply(report_row, function(x) (1/ex_rate) * x))  # EUR to CZK
    }
  }

  return(report_row)
}

get_report <- function(l) {
  df_final <- data.frame()
  for (i in 10:length(l[,1])) {
    if (substr(l[i,2], 1, 1) != 'T') {
      if (!substr(l[i,2], 1, 10) %in% acc_avoid) {
        df_final <- rbind(df_final, create_row(l, i))
      }
    }
  }
  colnames(df_final) <- as.character(names(read_excel(paste0(wd, 'data/FC_IC_Excel.xlsx'))))
  df_final <- df_final[, -c(16:(15+month))]
  df <- data.frame()
  for (i in seq(nrow(df_final))) {
    row <- df_final[i, 16:(ncol(df_final) - 4)]
    if (!all(is.na(row))) {
      df <- rbind(df, df_final[i,])
    }
  }
  return(df)
}


rep <- get_report(input)

# write.xlsx(rep, paste0(wd, '6925_czk.xlsx'))
