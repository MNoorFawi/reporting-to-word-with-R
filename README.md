reporting in word
================

Reporting Analysis done with R into a Word Document
---------------------------------------------------

### We will use **R** to analyse the [Bank Marketing](https://archive.ics.uci.edu/ml/datasets/Bank+Marketing) dataset and apply a **Logistic Regression** model on it, then deliver results in a **Word** document.

``` r
options(warn = -1)
library(officer)
library(flextable)
library(ggplot2)
library(reshape)
library(dplyr)
library(gridExtra)
library(e1071)
```

``` r
data <- read.csv("bank.csv", sep = ";")
str(data)
```

    ## 'data.frame':    4521 obs. of  17 variables:
    ##  $ age      : int  30 33 35 30 59 35 36 39 41 43 ...
    ##  $ job      : Factor w/ 12 levels "admin.","blue-collar",..: 11 8 5 5 2 5 7 10 3 8 ...
    ##  $ marital  : Factor w/ 3 levels "divorced","married",..: 2 2 3 2 2 3 2 2 2 2 ...
    ##  $ education: Factor w/ 4 levels "primary","secondary",..: 1 2 3 3 2 3 3 2 3 1 ...
    ##  $ default  : Factor w/ 2 levels "no","yes": 1 1 1 1 1 1 1 1 1 1 ...
    ##  $ balance  : int  1787 4789 1350 1476 0 747 307 147 221 -88 ...
    ##  $ housing  : Factor w/ 2 levels "no","yes": 1 2 2 2 2 1 2 2 2 2 ...
    ##  $ loan     : Factor w/ 2 levels "no","yes": 1 2 1 2 1 1 1 1 1 2 ...
    ##  $ contact  : Factor w/ 3 levels "cellular","telephone",..: 1 1 1 3 3 1 1 1 3 1 ...
    ##  $ day      : int  19 11 16 3 5 23 14 6 14 17 ...
    ##  $ month    : Factor w/ 12 levels "apr","aug","dec",..: 11 9 1 7 9 4 9 9 9 1 ...
    ##  $ duration : int  79 220 185 199 226 141 341 151 57 313 ...
    ##  $ campaign : int  1 1 1 4 1 2 1 2 2 1 ...
    ##  $ pdays    : int  -1 339 330 -1 -1 176 330 -1 -1 147 ...
    ##  $ previous : int  0 4 1 0 0 3 2 0 0 2 ...
    ##  $ poutcome : Factor w/ 4 levels "failure","other",..: 4 1 1 4 4 1 2 4 4 1 ...
    ##  $ y        : Factor w/ 2 levels "no","yes": 1 1 1 1 1 1 1 1 1 1 ...

``` r
summary(data)
```

    ##       age                 job          marital         education   
    ##  Min.   :19.00   management :969   divorced: 528   primary  : 678  
    ##  1st Qu.:33.00   blue-collar:946   married :2797   secondary:2306  
    ##  Median :39.00   technician :768   single  :1196   tertiary :1350  
    ##  Mean   :41.17   admin.     :478                   unknown  : 187  
    ##  3rd Qu.:49.00   services   :417                                   
    ##  Max.   :87.00   retired    :230                                   
    ##                  (Other)    :713                                   
    ##  default       balance      housing     loan           contact    
    ##  no :4445   Min.   :-3313   no :1962   no :3830   cellular :2896  
    ##  yes:  76   1st Qu.:   69   yes:2559   yes: 691   telephone: 301  
    ##             Median :  444                         unknown  :1324  
    ##             Mean   : 1423                                         
    ##             3rd Qu.: 1480                                         
    ##             Max.   :71188                                         
    ##                                                                   
    ##       day            month         duration       campaign     
    ##  Min.   : 1.00   may    :1398   Min.   :   4   Min.   : 1.000  
    ##  1st Qu.: 9.00   jul    : 706   1st Qu.: 104   1st Qu.: 1.000  
    ##  Median :16.00   aug    : 633   Median : 185   Median : 2.000  
    ##  Mean   :15.92   jun    : 531   Mean   : 264   Mean   : 2.794  
    ##  3rd Qu.:21.00   nov    : 389   3rd Qu.: 329   3rd Qu.: 3.000  
    ##  Max.   :31.00   apr    : 293   Max.   :3025   Max.   :50.000  
    ##                  (Other): 571                                  
    ##      pdays           previous          poutcome      y       
    ##  Min.   : -1.00   Min.   : 0.0000   failure: 490   no :4000  
    ##  1st Qu.: -1.00   1st Qu.: 0.0000   other  : 197   yes: 521  
    ##  Median : -1.00   Median : 0.0000   success: 129             
    ##  Mean   : 39.77   Mean   : 0.5426   unknown:3705             
    ##  3rd Qu.: -1.00   3rd Qu.: 0.0000                            
    ##  Max.   :871.00   Max.   :25.0000                            
    ## 

``` r
docx <- read_docx()
officer::body_add_par(docx, value = "Table of Content", 
                      style = "heading 1")
body_add_toc(docx, level = 2)
body_add_break(docx)
```

``` r
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
```

``` r
body_add_par(docx, "Exploring Data", style = "heading 1")
body_add_par(docx, "Summarizing Data", style = "heading 2")
body_add_par(docx, '- Examining the "age" and "balance" averages of each "Marital Status" group',
             style = "Normal")
age_balance_marital <- data %>% group_by(marital) %>%
  summarise(age_average = mean(age),
            balance_average = mean(balance))
abm <- create_flextable(age_balance_marital, "box")
flextable::body_add_flextable(docx, value = abm, align = "left")
```

``` r
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
```

``` r
body_add_par(docx, "", style = "Normal")
body_add_par(docx, '- Examining the "duration" of each "Contact Method"',
             style = "Normal")
contact_duration <- data %>% group_by(contact) %>%
  summarise(duration_average_min = mean(duration)) %>%
  mutate(duration_average_min = duration_average_min / 60)
cd <- create_flextable(contact_duration, "vanilla")
flextable::body_add_flextable(docx, value = cd, align = "left")
```

``` r
cursor_end(docx)
body_add_break(docx)
body_add_par(docx, "Visualizing Data", style = "heading 2")
body_add_par(docx, "- Visualizing distribution of numeric columns: ", 
             style = "Normal")
```

``` r
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
```

![](do_files/figure-markdown_github/visual-1.png)

``` r
filepath <- "."
filename <- paste0(dirname(filepath), "/density.png")
ggsave(filename = filename, height = 4, width = 6)
body_add_img(docx, src = filename, height = 4, width = 6)
```

``` r
body_add_par(docx, "", style = "Normal")
body_add_par(docx, "- Visualizing relationship between numeric columns and y: ", style = "Normal")
ggplot(age_balance_duration, aes(x = value, y = as.numeric(y))) +
  geom_point(position = position_jitter(w = 0.05, h = 0.05),
             alpha = 0.3) +
  geom_smooth(method = "loess", se = FALSE) +
  facet_wrap( ~ variable, scales = "free", ncol = 2) +
  theme_minimal() + ylab("Subscription")
```

![](do_files/figure-markdown_github/rel-1.png)

``` r
filename <- paste0(dirname(filepath), "/scatter.png")
ggsave(filename = filename, height = 4, width = 6)
body_add_img(docx, src = filename, height = 4, width = 6)
```
``` r
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
```
``` r
body_add_par(docx, "", style = "Normal")
body_add_par(docx, "- Visualizing frequency of categorical variables and their relationship with y: ", style = "Normal")

# function (not in)
'%!in%' <- function(x, y)
  ! ('%in%'(x, y))

vars <- colnames(data)
# summary of numeric data
summary(data[, vars[sapply(data[, vars], class) %in% 
                      c("numeric", "integer")]])
```

    ##       age           balance           day           duration   
    ##  Min.   :19.00   Min.   :-3313   Min.   : 1.00   Min.   :   4  
    ##  1st Qu.:33.00   1st Qu.:   69   1st Qu.: 9.00   1st Qu.: 104  
    ##  Median :39.00   Median :  444   Median :16.00   Median : 185  
    ##  Mean   :41.17   Mean   : 1423   Mean   :15.92   Mean   : 264  
    ##  3rd Qu.:49.00   3rd Qu.: 1480   3rd Qu.:21.00   3rd Qu.: 329  
    ##  Max.   :87.00   Max.   :71188   Max.   :31.00   Max.   :3025  
    ##     campaign          pdays           previous      
    ##  Min.   : 1.000   Min.   : -1.00   Min.   : 0.0000  
    ##  1st Qu.: 1.000   1st Qu.: -1.00   1st Qu.: 0.0000  
    ##  Median : 2.000   Median : -1.00   Median : 0.0000  
    ##  Mean   : 2.794   Mean   : 39.77   Mean   : 0.5426  
    ##  3rd Qu.: 3.000   3rd Qu.: -1.00   3rd Qu.: 0.0000  
    ##  Max.   :50.000   Max.   :871.00   Max.   :25.0000

``` r
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
```

![](do_files/figure-markdown_github/cats-1.png)

``` r
ggsave("catvars.png", height = 5, width = 7)
body_add_img(docx, src = "catvars.png", height = 5, width = 7)
```

``` r
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
```

``` r
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
```

``` r
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
```

![](do_files/figure-markdown_github/modeling-1.png)

``` r
ggsave("threshold.png", height = 4, width = 5)
body_add_img(docx, src = "threshold.png", height = 4, width = 5)
```

``` r
body_add_par(docx, "The model seems to classify well at probablity: 0.12",
             style = "Normal")
body_add_par(docx, "", style = "Normal")
body_add_par(docx, "Evaluating the Model", 
             style = "heading 2")
body_add_par(docx, "- Measuring Model Accuracy: ", style = "Normal")

(tab <- table(train$pred >= 0.12, train$y))
```

    ##        
    ##           no  yes
    ##   FALSE 2690   74
    ##   TRUE   508  344

``` r
(`train accuracy` <- round((tab[1, 1] + tab[2, 2]) / sum(tab), 2))
```

    ## [1] 0.84

``` r
test$pred <- predict(glmmodel, newdata = test, type = "response")
(tab <- table(test$pred >= 0.12, test$y))
```

    ##        
    ##          no yes
    ##   FALSE 683  15
    ##   TRUE  119  88

``` r
(`test accuracy` <- round((tab[1, 1] + tab[2, 2]) / sum(tab), 2))
```

    ## [1] 0.85

``` r
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
```

``` r
print(docx, target = "report.docx")
```
