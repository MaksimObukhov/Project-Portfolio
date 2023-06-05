library(dplyr)
library(tidyverse)

#load data
oscar_data <- read.csv("oscars.csv", encoding = "UTF-8")
movie_data <- read.csv("movies.csv", encoding = "UTF-8")

summary(oscar_data)
summary(movie_data["year"])

#subset awards that were given in the last 20 years and can be considered relevant
categories <- oscar_data[oscar_data["Year"] >= 1980,]

#we omit movies older than 1980 as there are none in the movies.csv dataset
oscar_data <- oscar_data[oscar_data["Year"] > 1979,]

# Remove missing film names
oscar_data <- oscar_data[oscar_data$Film != "", ]

#extract all unique awards a film can get
categories <- unique(categories["Award"])

#remove memorial awards
categories <- as.vector(categories %>% filter(!grepl("Award|Memorial|Medal", Award)))

#as factor
nominees <- oscar_data
nominees[is.na(nominees)] <- 0
nominees$Winner <- as.factor(nominees$Winner)

#filter only oscar winners from recent categories
nominees <- filter(nominees, Award %in% unlist(categories),)


colnames(nominees) <- tolower(colnames(nominees))
nominees <- nominees %>%
  mutate(across(where(is.character), str_trim))

nominees$year <- as.integer(nominees$year)

#join oscar nominee table with the film data
# we use inner join to keep relevant records
total_data <- inner_join(x=nominees , y=movie_data, by=c("name"="name", "year" = "year"), keep = FALSE)

dim(total_data)
winners <- total_data[total_data["winner"] == 1,]
dim(winners)


category_analysis <- function (category){
  par(mfrow=c(3,2))
  nom <- total_data
  win <- winners
  if (category != "All"){
  nom <- total_data[total_data["award"] == category,]
  win <- winners[winners["award"] == category,]
  }
  #visualize relative frequencies of top 6 countries of origins
  barplot(tail(sort(prop.table(table(nom["country"])))),
          main = paste("Top 6 countries by relative frequency of oscar nominations for ", category),
          xlab = "Country name",
          ylab = "Relative frequency of nominations")

  #visualize relative frequencies of top 6 countries of origins
  barplot(tail(sort(prop.table(table(win["country"])))),
          main = paste("Top 6 countries by relative frequency of oscar wins for ", category),
          xlab = "Country name",
          ylab = "Relative frequency of wins")

  #visualize relative frequencies of top 6 production companies
  barplot(tail(sort(prop.table(table(nom["company"])))),
          main = paste("Top 6 producion companies by relative frequency of oscar nominations for ", category),
          xlab = "Company name",
          ylab = "Relative frequency of nominations")

  barplot(tail(sort(prop.table(table(win["company"])))),
          main = paste("Top 6 production companies by relative frequency of oscar wins for ", category),
          xlab = "Company name",
          ylab = "Relative frequency of wins")

  barplot(tail(sort(prop.table(table(nom["director"])))),
          main = paste("Top 6 directors by relative frequency of oscar nominations for ", category),
          xlab = "Director's name",
          ylab = "Relative frequency of wins")

  barplot(tail(sort(prop.table(table(win["director"])))),
          main = paste("Top 6 directors by relative frequency of oscar wins for ", category),
          xlab = "Director's name",
          ylab = "Relative frequency of wins")

}

category_analysis("Best Picture")
category_analysis("Directing")
category_analysis("All")

# scores on winning films and nominated films
total_data %>% ggplot(aes(x=score, group = winner, fill = winner)) +
  geom_density(alpha = 0.7) +
  ggtitle("Density plot of scores of winning films (1) and nominated films (0)")
d <-
  total_data %>%
  filter(!is.na(score), !is.na(winner)) |>
  select(name, year, score, winner) |>
  distinct()
score_test <- wilcox.test(score ~ winner, data = d)
d %>%
  ggplot(aes(x=winner, y=score, fill=winner)) +
  geom_boxplot() +
  labs(title = sprintf("Score is statistically significant (p-val %.4f)",
                       score_test$p.value),
       subtitle = "Test used: Wilcoxon two-paired test",
       x = "Winner", y = "Score of a movie") +
  theme_bw()

# runtimes of winning films and nominated films
total_data %>% ggplot(aes(x=runtime, group=winner, fill=winner)) +
  geom_density(alpha = 0.7) +
  ggtitle("Density plot of runtimes of winning films (1) and nominated films (0)")
# budget and gross of a film, colors by nominee/winner
d <-
  total_data %>%
  filter(!is.na(runtime), !is.na(winner)) |>
  select(name, year, runtime, winner) |>
  distinct()
runtime_test <- t.test(runtime ~ winner, data = d)
d |>
  ggplot(aes(x=winner, y=runtime, fill=winner))+
  geom_boxplot() +
  labs(title = sprintf("Runtime is statistically significant (p-val %.4f)",
                       runtime_test$p.value),
       subtitle = "Test used: Two-sample t-test; One missing runtime removed",
       x = "Winner", y = "Runtime of a movie") +
  theme_bw()

# budgets of winning and nominated films
total_data %>% ggplot(aes(x=budget, group=winner, fill=winner)) +
  geom_density(alpha = 0.7) +
  ggtitle("Density plot of budgets of winning films (1) and nominated films(0)")


# budget and gross of a film, colors by nominee/winner
d <-
  total_data %>%
  filter(!is.na(budget), !is.na(gross)) |>
  select(name, year, budget, gross) |>
  distinct()
budget_corr <- cor(d$budget, d$gross)
d |>
  ggplot(aes(x=budget, y=gross)) +
  geom_point() +
  geom_label(data = filter(d, gross > 2000000000), aes(label = name),
             nudge_x = -30000000) +
  geom_label(data = filter(d, budget > 250000000), aes(label = name),
             nudge_y = 200000000, nudge_x = 70000000) +
  geom_smooth(formula = y ~ x, method = "lm", se = F)  +
  scale_x_continuous(labels = scales::dollar_format(), limits = c(0, 4e8)) +
  scale_y_continuous(labels = scales::dollar_format()) +
  labs(title = "Budget and Gross earnings are lineary correlated",
       subtitle = sprintf("Pearsons correlation: %.2f", budget_corr),
       x = "Budget", y = "Gross earning") +
  theme_bw()
total_data %>% ggplot(aes(x=winner, y=budget, fill=winner)) +
  geom_boxplot() +
  ggtitle("Boxplots of budgets")
total_data %>% ggplot(aes(x=winner, y=gross, fill=winner)) +
  geom_boxplot() +
  ggtitle("Boxplots of gross")

# votes vs score
total_data %>% ggplot(aes(x=votes, y=score)) +
  geom_point() +
  ggtitle("Scatter plots of votes and scores")

# runtime vs budget
total_data %>% ggplot(aes(x=runtime, y=budget, col=winner)) +
geom_point() +
  ggtitle("Scatter plots of runtimes and budgets")

# year vs score
total_data %>% ggplot(aes(x=year, y=score, col=winner)) +
  geom_point() +
  ggtitle("Scatter plots of years and scores")

# gross vs score
total_data %>% ggplot(aes(x=gross, y=score)) +
  geom_point() +
  ggtitle("Scatter plots of gross and scores") # possible outliers?

# budget vs score
total_data %>% ggplot(aes(x=budget, y=score)) +
  geom_point() +
  ggtitle("Scatter plots of gross and scores")
