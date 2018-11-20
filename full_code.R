
## Reporting Analysis done with R into a Word Document

options(warn = -1)
library(officer)
library(flextable)
library(ggplot2)
library(reshape)
library(dplyr)
library(gridExtra)
library(e1071)

data <- read.csv("bank.csv", sep = ";")
str(data)
summary(data)

docx <- read_docx()
officer::body_add_par(docx, value = "Table of Content", 
                      style = "heading 1")
body_add_toc(docx, level = 2)
body_add_break(docx)

title_case <- function(x) {
  lower_case <- tolower(x)
  without_underscore <- gsub("_", " ", lower_case, perl = TRUE)
  titled <- gsub("(^|[[:space:]])([[:alpha:]])",
                 "\\1\\U\\2",
                 without_underscore,
                 perl = TRUE)
  titled
}

# trim <- function(x) {
#   gsub("^\\s+|\\s+$", "", x)
# }

create_flextable <- function(df, theme = "box") {
  thm <- paste0("flextable::theme_", theme, "(ft)")
  ft <- flextable::flextable(df)
  ns <- title_case(names(df))
  ex <- paste0('"', names(df), '"', "=", '"', ns, '"', collapse = ",")
  expr <- paste0("flextable::set_header_labels(ft,")
  exp <- paste0(expr, ex, ")")
  ft <- eval(parse(text = exp))
  ft <- eval(parse(text = thm))
  ft <- ft %>%
    flextable::bold(part = "header") %>%
    flextable::bold(j = 1) %>%
    flextable::font(fontname = "Arrial Narrow", part = "all") %>%
    flextable::fontsize(size = 10, part = "all") %>%
    flextable::align(align = "left", part = "all") %>%
    flextable::autofit()
  ft
}

body_add_par(docx, "Exploring Data", style = "heading 1")
body_add_par(docx, "Summarizing Data", style = "heading 2")
body_add_par(docx, '- Examining the "age" and "balance" averages of each "Marital Status" group',
             style = "Normal")
age_balance_marital <- data %>% group_by(marital) %>%
  summarise(age_average = mean(age),
            balance_average = mean(balance))
abm <- create_flextable(age_balance_marital, "box")
flextable::body_add_flextable(docx, value = abm, align = "left")

body_add_par(docx, "", style = "Normal")
body_add_par(docx, '- Examining the "age" and "balance" averages of each "Job" who subscribed in this campaign',
             style = "Normal")
age_balance_job <-
  data %>% group_by(job, y) %>% filter(y == "yes") %>%
  summarise(age_average = mean(age),
            balance_average = mean(balance)) %>%
  arrange(desc(balance_average))
abj <- create_flextable(age_balance_job, "zebra")
flextable::body_add_flextable(docx, value = abj, align = "left")

body_add_par(docx, "", style = "Normal")
body_add_par(docx, '- Examining the "duration" of each "Contact Method"',
             style = "Normal")
contact_duration <- data %>% group_by(contact) %>%
  summarise(duration_average_min = mean(duration)) %>%
  mutate(duration_average_min = duration_average_min / 60)
cd <- create_flextable(contact_duration, "vanilla")
flextable::body_add_flextable(docx, value = cd, align = "left")

cursor_end(docx)
body_add_break(docx)
body_add_par(docx, "Visualizing Data", style = "heading 2")
body_add_par(docx, "- Visualizing distribution of numeric columns: ", 
             style = "Normal")

# in case we don't know the exact spelling of the column names we can 
# extract them with this regex
regex <- "(?=.*age)|(?=.*balanc)|(?=.*durat)"
cols <- colnames(data)
nums <- grep(regex, cols, perl = TRUE, value = TRUE)

age_balance_duration <- data[, c(nums, "y")] %>%
  mutate(balance = log(balance, base = 10),
         duration = duration / 60) %>%
  dplyr::rename(log10_balance = balance,
                duration_min = duration) %>%
  melt(id.vars = "y")

ggplot(age_balance_duration, aes(x = value, fill = y)) +
  geom_density(color = "white", alpha = 0.6) +
  facet_wrap( ~ variable, scales = "free") +
  theme_minimal() + scale_fill_brewer(palette = "Set1") +
  theme(legend.position = "top")

filepath <- "."
filename <- paste0(dirname(filepath), "/density.png")
ggsave(filename = filename, height = 4, width = 6)
body_add_img(docx, src = filename, height = 4, width = 6)

body_add_par(docx, "", style = "Normal")
body_add_par(docx, "- Visualizing relationship between numeric columns and y: ", style = "Normal")
ggplot(age_balance_duration, aes(x = value, y = as.numeric(y))) +
  geom_point(position = position_jitter(w = 0.05, h = 0.05),
             alpha = 0.3) +
  geom_smooth(method = "loess", se = FALSE) +
  facet_wrap( ~ variable, scales = "free", ncol = 2) +
  theme_minimal() + ylab("Subscription")

filename <- paste0(dirname(filepath), "/scatter.png")
ggsave(filename = filename, height = 4, width = 6)
body_add_img(docx, src = filename, height = 4, width = 6)

body_add_par(docx, "", style = "Normal")
body_add_par(docx, "- spotting days, months and jobs in which users tend to subscribe", style = "Normal")

data <- data %>%
  mutate(month =
           factor(gsub("(^[[:alpha:]])", "\\U\\1",
                       month, perl = TRUE),
                  month.abb))
gg <- ggplot(data, aes(x = as.numeric(month), color = y)) +
  geom_freqpoly(stat = "count") +
  xlab("month") + 
  scale_x_continuous(breaks = seq(1, 12, 1), labels = month.abb) +
  scale_y_continuous(breaks = seq(0, 1500, 250)) +
  theme_minimal() + scale_color_brewer(palette = "Set1") +
  xlab("last contact month of the year")

gg2 <- ggplot(data, aes(x = job, fill = y)) +
  geom_bar(width = 0.8,
           color = "white",
           alpha = 0.8) +
  theme_minimal() + scale_fill_brewer(palette = "Set1")

gg3 <- ggplot(data, aes(x = day, fill = y)) +
  geom_histogram(bins = 30, col = "white") + theme_minimal() +
  scale_fill_brewer(palette = "Accent") +
  xlab("last contact day of the month")

ggsave("daymonthjob.png", 
       arrangeGrob(gg, gg2, gg3), height = 4, width = 7)
body_add_img(docx, src = "daymonthjob.png", height = 4, width = 7)

body_add_par(docx, "", style = "Normal")
body_add_par(docx, "- Visualizing frequency of categorical variables and their relationship with y: ", style = "Normal")

# function (not in)
'%!in%' <- function(x, y)
  ! ('%in%'(x, y))

vars <- colnames(data)
# summary of numeric data
summary(data[, vars[sapply(data[, vars], class) %in% 
                      c("numeric", "integer")]])

# change pdays variable to factor of intervals and delete it
data$days_after_last_contact <-
  cut(data$pdays,
      breaks = c(-1, 0, 150, seq(300, 900, 300)),
      include.lowest = T)
data$pdays <- NULL

# change campaign and previous to factor of intervals and rename them
data <- data %>% mutate(
  campaign = cut(
    campaign,
    breaks = c(seq(1, 10, 2),
               seq(25, 50, 25)),
    include.lowest = TRUE
  ),
  previous = cut(
    previous,
    breaks = c(seq(0, 6, 2),
               seq(15, 25, 10)),
    include.lowest = TRUE
  )
) %>% dplyr::rename(
  number_contacts_this_campaign = campaign,
  number_contacts_before_this_campaign = previous
)

# new column names
vars <- colnames(data)
# categorical variables
cat_vars <-
  vars[sapply(data[, vars], class) %in% c("factor", "character")]
# we have already examined month and job
cat_vars <- cat_vars[cat_vars %!in% c("month", "job")]

cat_melted <- data[, cat_vars] %>%
  melt(id.vars = "y")
ggplot(cat_melted, aes(x = value, fill = y)) + 
  geom_bar(width = 0.5, color = "white") +
  theme_minimal() + theme(
    legend.position = c(0.6, 0.1),
    legend.direction = "horizontal",
    legend.key.size = unit(10, "mm"),
    legend.title = element_text("Subscription")
  ) +
  facet_wrap(~ variable, scales = "free", ncol = 3) +
  scale_fill_brewer(palette = "Paired")

ggsave("catvars.png", height = 5, width = 7)
body_add_img(docx, src = "catvars.png", height = 5, width = 7)

body_add_break(docx)
body_add_par(docx, "Modeling Data", style = "heading 1")
body_add_par(docx, "Preparing data for modeling", 
             style = "heading 2")
body_add_par(docx, "- one hot encoding of categorical variables and normalizing numeric ones", style = "Normal")
body_add_par(docx, paste("dimenstions of original data:",
                         paste(dim(data), collapse = " , ")),
             style = "Normal")

data$day <- factor(data$day)
# predictors and outcome
pred_vars <- setdiff(vars, "y")

## to one hot encode factor values and normalize numeric ones if needed
cat <-
  pred_vars[sapply(data[, pred_vars], class) %in% c("factor", "character")]
num <-
  pred_vars[sapply(data[, pred_vars], class) %in% c("numeric", "integer")]

for (i in cat) {
  dict <- unique(data[, i])
  for (key in dict) {
    data[[paste0(i, "_", key)]] <- 1.0 * (data[, i] == key)
  }
}

data[, num] <- apply(data[, num], 2, function(x) {
  (x - min(x)) / (max(x) - min(x))
})

data <- data[, -which(colnames(data) %in% cat)]

body_add_par(docx, paste("dimenstions of new encoded data:",
                         paste(dim(data), collapse = " , ")),
             style = "Normal")

cursor_end(docx)
body_add_par(docx, "Train/Test Splitting", 
             style = "heading 2")
body_add_par(docx, "- splitting data to train and test datasets (80/20)", style = "Normal")

sample_size <- floor(0.8 * nrow(data))
## set the seed to make your splits reproducible
set.seed(13)
train_indices <- sample(seq(nrow(data)), size = sample_size)
train <- data[train_indices,]
test <- data[-train_indices,]

body_add_par(docx, paste("nrows of training dataset:", 
                         nrow(train)), 
             style = "Normal")
body_add_par(docx, paste("nrows of test dataset:", 
                         nrow(test)), 
             style = "Normal")

body_add_par(docx, "Fitting Classification Model", 
             style = "heading 2")
body_add_par(docx, "- Applying Logistic Regression algorithm to data: ", style = "Normal")

glmmodel <- glm(y ~ ., data = train, family = binomial(link = "logit"))

# how well the model fits the data it has seen
train$pred <- predict(glmmodel, newdata = train, type = "response")

# visualizing classification thresheold
body_add_par(docx, "", style = "Normal")
body_add_par(docx, "Spotting Classification Threshold Probability Value: ",
             style = "Normal")
ggplot(train, aes(x = pred, fill = y)) + 
  geom_density(alpha = 0.6, col = "white") +
  scale_fill_brewer(palette = "Set1") + theme_minimal()
ggsave("threshold.png", height = 4, width = 5)
body_add_img(docx, src = "threshold.png", height = 4, width = 5)
body_add_par(docx, "The model seems to classify well at probablity: 0.12",
             style = "Normal")

body_add_par(docx, "", style = "Normal")
body_add_par(docx, "Evaluating the Model", 
             style = "heading 2")
body_add_par(docx, "- Measuring Model Accuracy: ", style = "Normal")

(tab <- table(train$pred >= 0.12, train$y))
(`train accuracy` <- round((tab[1, 1] + tab[2, 2]) / sum(tab), 2))
test$pred <- predict(glmmodel, newdata = test, type = "response")
(tab <- table(test$pred >= 0.12, test$y))
(`test accuracy` <- round((tab[1, 1] + tab[2, 2]) / sum(tab), 2))
accuracy <- t(cbind(`train accuracy`, `test accuracy`))
accuracy <- as.data.frame(cbind(rownames(accuracy), accuracy))

ns <- names(accuracy)
ft <- flextable(accuracy)
# to exclude headers in flextable
# set_header_labels(ft, V1 = NULL, V2 = NULL)
# names(df) <- paste0("X", 1:ncol(df))

# so we need to know in advance the column names but there's a trick 
# to apply it without knowing them
ex <- paste0('"', ns, '"', "=NULL", collapse = ",")
expr <- paste0('flextable::set_header_labels(ft,')
exp <- paste0(expr, ex, ")")
ft <- eval(parse(text = exp))

ft <- ft %>% bold(j = 1) %>% font(fontname = "Arial") %>%
  fontsize(size = 11) %>% border_remove() %>% autofit() %>%
  align(align = "left", part = "all") %>% 
  color(i = 2, j = 2, color = "darkblue")
flextable::body_add_flextable(docx, value = ft, align = "left")

print(docx, target = "report.docx")

