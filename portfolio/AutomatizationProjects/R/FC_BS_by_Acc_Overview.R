library(tidyverse)
library(stringi)
library(readxl)
library(openxlsx)


input_name <- list.files(path = 'report/data/input/FC_BS/', pattern = '*.xlsx')
input_bs <- as.data.frame(read_excel(paste0('report/data/input/FC_BS/', input_name[1]), col_names = F))
input_pl <- as.data.frame(read_excel(paste0('report/data/input/FC_BS/', input_name[2]), col_names = F))

map_ud2 <- read_excel('bs_overview/data/Map_Accounts_KPIs_ReCo_to_OS_20220413.xlsx')[-1,c(3,5)]
ud2_shorts <- map_ud2[map_ud2[2] == 'mt_short',][1]


date <- '2022M3'
year <- as.numeric(substr(date, 1, 4))
month <- as.numeric(substr(date, 6, 6))
czk_eur <- 0.040570206


ent_curr <- read_excel('report/data/entity_currency.xlsx')
acc_avoid <- c('SDD0000001', 'AF08000001')

eur_ent <- c()
for (i in seq(nrow(ent_curr))) {
  if (ent_curr[i, 2] == 'EUR') {
    eur_ent <- as.character(append(eur_ent, ent_curr[i, 1]))
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

sc <- get_sc(input_bs[4,1])

create_row <- function(l, i) {
  entity <- strsplit(l[6,1], ' ')[[1]][2]
  acc <- substr(l[i,1], 1, 10)

  df <- data.frame(
    E = entity,
    S.C = sc,
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
    TV_Ann = NA,
    V = 'YTD')

  row <- get_row(l, i, entity, acc_first_letter(acc), czk_eur)
  if (!is.null(row)) {
    df <- cbind(df, row) %>%
    cbind(data.frame(NA, 'IN61', NA, 'Import', NA))
    return(df)
  }

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


get_row <- function(l, i, entity, acc_fl, ex_rate) {
  options(digits=15)
  report_row <- l[i, 2:ncol(l)]

  colnames(report_row) <- lapply(colnames(report_row), function(x) {
    gsub(' ', '.', x)
  })

  report_row <- report_row %>%
    mutate_all(~replace(., . == '0', NA)) %>%
    mutate_all(~as.numeric(.))

  if (all(is.na(report_row))) {
    return (NULL)
  }

  if (acc_fl %in% c('R', 'C', 'S')) {
    report_row <- as.data.frame(lapply(report_row, function(x) -1 * x))
  }

  # if (entity %in% eur_ent) {
  #   report_row <- as.data.frame(lapply(report_row, function(x) (1/ex_rate) * x))
  # }

  return(report_row)
}

get_report <- function(l) {
  df_final <- data.frame()
  for (i in 9:nrow(l)) {
    if (l[i,1] != 'EK01000003') {
      df_final <- rbind(df_final, create_row(l, i))
    }
  }

  colnames(df_final) <- as.character(names(read_excel('report/data/FC_IC_Excel.xlsx')))
  df_final <- df_final[, -c(16:(15+month))]
  df <- data.frame()
  for (i in seq(nrow(df_final))) {
    row <- df_final[i, 16:(ncol(df_final) - 5)]
    if (!all(is.na(row))) {
      row <- row %>%
        mutate_all(~replace(., is.na(.), 0)) %>%
        mutate_all(~as.numeric(.))

      df_final[i, 16:(ncol(df_final) - 5)] <- row
      df <- rbind(df, df_final[i,])
    }
  }
  return(df)
}


bs <- get_report(input_bs)
pl <- get_report(input_pl)

res  <- rbind(bs, pl)

# write.xlsx(res, 'report.xlsx')

