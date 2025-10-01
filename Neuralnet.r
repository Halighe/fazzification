# Установка пакетов (если не установлены)
if (!require("openxlsx")) install.packages("openxlsx")
if (!require("neuralnet")) install.packages("neuralnet")
if (!require("NeuralNetTools")) install.packages("NeuralNetTools")
if (!require("readxl")) install.packages("readxl")

library(readxl)
library(openxlsx)
library(neuralnet)
library(NeuralNetTools)

# Функции
f <- function(a, b, c, fuz, koef){
  for(i in 1:length(fuz[, 1])){
    if(a < fuz[i, 1] & fuz[i, 1] <= c) {fuz[i, 1] <- round(((fuz[i, 1] - a) / (c - a)), 2) + koef}
    else if(c < fuz[i, 1] & fuz[i, 1] < b) {fuz[i, 1] <- round(((b - fuz[i, 1]) / (b - c)), 2) + koef}
    else {fuz[i, 1] <- 0 + koef}
  }
  return(fuz)
}

RMSE <- function(a, b, c){
  arr <- array(0, dim=c)
  for(i in 1:c){ 
    arr[i] <- sqrt(mean((a[i] - b[i])^2))
  }
  return(arr)
}

# Функция фазификации для интенсивности движения (3 фазы)
fuzz_intensity <- function(data, numcol){
  fuzz_low <- array(0, dim = length(data[,numcol]))
  fuzz_medium <- array(0, dim = length(data[,numcol]))
  fuzz_high <- array(0, dim = length(data[,numcol]))
  
  for(f in 1:length(data[,numcol])){
    if(data[f,numcol] >= 0 & data[f,numcol] < 750){
      fuzz_low[f] <- 1
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] >= 750 & data[f,numcol] < 900){
      fuzz_medium[f] <- 1
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] >= 900 & data[f,numcol] <= 2000){
      fuzz_high[f] <- 1
      data[f,numcol] <- 0
    }
    else {
      data[f,numcol] <- 0
    }
  }
  c <- cbind(data[, numcol], fuzz_low, fuzz_medium, fuzz_high)
  return(c)
}

# Функция фазификации для времени горения зеленого (3 фазы)
fuzz_green_time <- function(data, numcol){
  fuzz_low <- array(0, dim = length(data[,numcol]))
  fuzz_medium <- array(0, dim = length(data[,numcol]))
  fuzz_high <- array(0, dim = length(data[,numcol]))
  
  for(f in 1:length(data[,numcol])){
    if(data[f,numcol] >= 0 & data[f,numcol] < 30){
      fuzz_low[f] <- 1
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] >= 30 & data[f,numcol] < 45){
      fuzz_medium[f] <- 1
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] >= 45 & data[f,numcol] <= 120){
      fuzz_high[f] <- 1
      data[f,numcol] <- 0
    }
    else {
      data[f,numcol] <- 0
    }
  }
  c <- cbind(data[, numcol], fuzz_low, fuzz_medium, fuzz_high)
  return(c)
}

# Функция преобразования в вектор
vec <- function(data, coff, col){
  p <- array(0, length(data[,1]))
  k <- 0
  for(i in 1:col){
    for(j in 1:length(data[,1])){
      if(data[j,i+coff] == 1){ p[j] <- 1 + k }
    }
    k <- k + 1
  }
  return(p)
}

# Функции для генерации комплексных правил
generate_comprehensive_rule <- function(fuzzy_row, original_row, row_index) {
  tryCatch({
    conditions <- c()
    
    # ПЕРЕКРЕСТОК 1
    if(fuzzy_row['Intensity_In_1_low'] == 1) conditions <- c(conditions, "вход1=низкий")
    if(fuzzy_row['Intensity_In_1_medium'] == 1) conditions <- c(conditions, "вход1=средний")
    if(fuzzy_row['Intensity_In_1_high'] == 1) conditions <- c(conditions, "вход1=высокий")
    
    if(fuzzy_row['Intensity_Straight_1_low'] == 1) conditions <- c(conditions, "прямо1=низкое")
    if(fuzzy_row['Intensity_Straight_1_medium'] == 1) conditions <- c(conditions, "прямо1=среднее")
    if(fuzzy_row['Intensity_Straight_1_high'] == 1) conditions <- c(conditions, "прямо1=высокое")
    
    if(fuzzy_row['Intensity_Right_1_low'] == 1) conditions <- c(conditions, "направо1=низкое")
    if(fuzzy_row['Intensity_Right_1_medium'] == 1) conditions <- c(conditions, "направо1=среднее")
    if(fuzzy_row['Intensity_Right_1_high'] == 1) conditions <- c(conditions, "направо1=высокое")
    
    if(fuzzy_row['Green_Time_1_low'] == 1) conditions <- c(conditions, "зеленый1=короткий")
    if(fuzzy_row['Green_Time_1_medium'] == 1) conditions <- c(conditions, "зеленый1=средний")
    if(fuzzy_row['Green_Time_1_high'] == 1) conditions <- c(conditions, "зеленый1=длинный")
    
    # ПЕРЕКРЕСТОК 2
    if(fuzzy_row['Intensity_In_2_low'] == 1) conditions <- c(conditions, "вход2=низкий")
    if(fuzzy_row['Intensity_In_2_medium'] == 1) conditions <- c(conditions, "вход2=средний")
    if(fuzzy_row['Intensity_In_2_high'] == 1) conditions <- c(conditions, "вход2=высокий")
    
    if(fuzzy_row['Intensity_Straight_2_low'] == 1) conditions <- c(conditions, "прямо2=низкое")
    if(fuzzy_row['Intensity_Straight_2_medium'] == 1) conditions <- c(conditions, "прямо2=среднее")
    if(fuzzy_row['Intensity_Straight_2_high'] == 1) conditions <- c(conditions, "прямо2=высокое")
    
    if(fuzzy_row['Intensity_Right_2_low'] == 1) conditions <- c(conditions, "направо2=низкое")
    if(fuzzy_row['Intensity_Right_2_medium'] == 1) conditions <- c(conditions, "направо2=среднее")
    if(fuzzy_row['Intensity_Right_2_high'] == 1) conditions <- c(conditions, "направо2=высокое")
    
    if(fuzzy_row['Green_Time_2_low'] == 1) conditions <- c(conditions, "зеленый2=короткий")
    if(fuzzy_row['Green_Time_2_medium'] == 1) conditions <- c(conditions, "зеленый2=средний")
    if(fuzzy_row['Green_Time_2_high'] == 1) conditions <- c(conditions, "зеленый2=длинный")
    
    # ЗАКЛЮЧЕНИЕ (ОБЩАЯ НАГРУЗКА)
    conclusion <- ""
    if(fuzzy_row['Total_Intensity_low'] == 1) conclusion <- "нагрузка=низкая"
    if(fuzzy_row['Total_Intensity_medium'] == 1) conclusion <- "нагрузка=средняя"
    if(fuzzy_row['Total_Intensity_high'] == 1) conclusion <- "нагрузка=высокая"
    
    # ФОРМИРУЕМ ПРАВИЛО
    if(length(conditions) > 0 & conclusion != "") {
      rule <- paste("ЕСЛИ", paste(conditions, collapse = " И "), 
                    "ТО", conclusion)
      return(rule)
    } else {
      return(paste("Маршрут", row_index, ": Неполные данные для формирования правила"))
    }
    
  }, error = function(e) {
    return(paste("Ошибка в маршруте", row_index, ":", e$message))
  })
}

# Функция для человеко-понятных правил
generate_human_readable_rule <- function(fuzzy_row, original_row, row_index) {
  tryCatch({
    conditions <- c()
    
    # Перекресток 1 - более понятные формулировки
    if(fuzzy_row['Intensity_In_1_high'] == 1) conditions <- c(conditions, "на_входе_1_много_машин")
    if(fuzzy_row['Intensity_Straight_1_high'] == 1) conditions <- c(conditions, "прямо_1_сильная_нагрузка")
    if(fuzzy_row['Intensity_Right_1_high'] == 1) conditions <- c(conditions, "направо_1_интенсивный_поток")
    if(fuzzy_row['Green_Time_1_low'] == 1) conditions <- c(conditions, "зеленый_1_короткий")
    
    if(fuzzy_row['Intensity_In_1_medium'] == 1) conditions <- c(conditions, "на_входе_1_средняя_нагрузка")
    if(fuzzy_row['Intensity_Straight_1_medium'] == 1) conditions <- c(conditions, "прямо_1_умеренная_нагрузка")
    if(fuzzy_row['Green_Time_1_medium'] == 1) conditions <- c(conditions, "зеленый_1_средний")
    
    if(fuzzy_row['Intensity_In_1_low'] == 1) conditions <- c(conditions, "на_входе_1_мало_машин")
    if(fuzzy_row['Intensity_Straight_1_low'] == 1) conditions <- c(conditions, "прямо_1_слабая_нагрузка")
    if(fuzzy_row['Green_Time_1_high'] == 1) conditions <- c(conditions, "зеленый_1_длинный")
    
    # Перекресток 2
    if(fuzzy_row['Intensity_In_2_high'] == 1) conditions <- c(conditions, "на_входе_2_много_машин")
    if(fuzzy_row['Intensity_Straight_2_high'] == 1) conditions <- c(conditions, "прямо_2_сильная_нагрузка")
    if(fuzzy_row['Intensity_Right_2_high'] == 1) conditions <- c(conditions, "направо_2_интенсивный_поток")
    if(fuzzy_row['Green_Time_2_low'] == 1) conditions <- c(conditions, "зеленый_2_короткий")
    
    if(fuzzy_row['Intensity_In_2_medium'] == 1) conditions <- c(conditions, "на_входе_2_средняя_нагрузка")
    if(fuzzy_row['Intensity_Straight_2_medium'] == 1) conditions <- c(conditions, "прямо_2_умеренная_нагрузка")
    if(fuzzy_row['Green_Time_2_medium'] == 1) conditions <- c(conditions, "зеленый_2_средний")
    
    if(fuzzy_row['Intensity_In_2_low'] == 1) conditions <- c(conditions, "на_входе_2_мало_машин")
    if(fuzzy_row['Intensity_Straight_2_low'] == 1) conditions <- c(conditions, "прямо_2_слабая_нагрузка")
    if(fuzzy_row['Green_Time_2_high'] == 1) conditions <- c(conditions, "зеленый_2_длинный")
    
    # Заключение
    conclusion <- ""
    if(fuzzy_row['Total_Intensity_high'] == 1) conclusion <- "общая_нагрузка=высокая"
    if(fuzzy_row['Total_Intensity_medium'] == 1) conclusion <- "общая_нагрузка=средняя"
    if(fuzzy_row['Total_Intensity_low'] == 1) conclusion <- "общая_нагрузка=низкая"
    
    if(length(conditions) > 0 & conclusion != "") {
      rule <- paste("Маршрут", row_index, ": ЕСЛИ", paste(conditions, collapse = " И "), 
                    "ТО", conclusion)
      return(rule)
    } else {
      return(paste("Маршрут", row_index, ": недостаточно_данных"))
    }
    
  }, error = function(e) {
    return(paste("Ошибка:", e$message))
  })
}

# Функция для приоритетных правил
generate_priority_rule <- function(fuzzy_row, original_row, row_index) {
  tryCatch({
    # Выбираем только наиболее значимые условия (не более 6)
    priority_conditions <- c()
    
    # Приоритет 1: Высокие интенсивности
    if(fuzzy_row['Intensity_In_1_high'] == 1) priority_conditions <- c(priority_conditions, "вход1=высокий")
    if(fuzzy_row['Intensity_In_2_high'] == 1) priority_conditions <- c(priority_conditions, "вход2=высокий")
    
    # Приоритет 2: Короткие времена зеленого
    if(fuzzy_row['Green_Time_1_low'] == 1) priority_conditions <- c(priority_conditions, "зеленый1=короткий")
    if(fuzzy_row['Green_Time_2_low'] == 1) priority_conditions <- c(priority_conditions, "зеленый2=короткий")
    
    # Приоритет 3: Высокие прямые потоки
    if(fuzzy_row['Intensity_Straight_1_high'] == 1 && length(priority_conditions) < 6) {
      priority_conditions <- c(priority_conditions, "прямо1=высокое")
    }
    if(fuzzy_row['Intensity_Straight_2_high'] == 1 && length(priority_conditions) < 6) {
      priority_conditions <- c(priority_conditions, "прямо2=высокое")
    }
    
    # Приоритет 4: Средние значения (если есть место)
    if(length(priority_conditions) < 4) {
      if(fuzzy_row['Intensity_In_1_medium'] == 1) priority_conditions <- c(priority_conditions, "вход1=средний")
      if(fuzzy_row['Intensity_In_2_medium'] == 1) priority_conditions <- c(priority_conditions, "вход2=средний")
    }
    
    # Заключение
    conclusion <- ""
    if(fuzzy_row['Total_Intensity_high'] == 1) conclusion <- "нагрузка=высокая"
    if(fuzzy_row['Total_Intensity_medium'] == 1) conclusion <- "нагрузка=средняя"
    if(fuzzy_row['Total_Intensity_low'] == 1) conclusion <- "нагрузка=низкая"
    
    if(length(priority_conditions) > 0 & conclusion != "") {
      rule <- paste("ЕСЛИ", paste(priority_conditions, collapse = " И "), 
                    "ТО", conclusion)
      return(rule)
    } else {
      return(paste("Маршрут", row_index, ": невозможно_сформировать_правило"))
    }
    
  }, error = function(e) {
    return(paste("Ошибка:", e$message))
  })
}

# Функция формирования правил из нейросети (для обратной совместимости)
generate_rules <- function(rez, target_var) {
  tryCatch({
    w <- garson(rez)
    pr_sort <- table(w$data[order(w$data[,1]),])
    
    sr <- array(0, dim = min(8, ncol(pr_sort)))
    x <- max(grep(target_var, colnames(pr_sort)))
    
    for(j in 1:length(sr)){
      if(x > 0 && x <= ncol(pr_sort)) {
        punkt <- colnames(pr_sort)[x]
        symblos <- substr(punkt, start = 1, stop = 5)
        c <- grep(symblos, colnames(pr_sort))
        if(length(c) > 0) {
          pr_sort <- pr_sort[-c, -c, drop = FALSE]
        }
        
        num_s <- gregexpr("_", punkt)[[1]][1] 
        if(num_s > 0) {
          substr(punkt, start = num_s, stop = num_s) <- "="
        }
        sr[j] <- punkt
        if(j < length(sr) && ncol(pr_sort) > 0) {
          x <- ncol(pr_sort)
        }
      }
    }
    
    pr <- "если:"
    if(length(sr) > 1) {
      for(i in 2:length(sr)){
        if(sr[i] != "") {
          pr <- paste(pr, sr[i]) 
        }
      }
      pr <- paste(pr, "то:", sr[1])
    }
    return(pr)
  }, error = function(e) {
    return(paste("Ошибка при формировании правил:", e$message))
  })
}

# ОСНОВНОЙ КОД

# ЗАГРУЗКА ДАННЫХ
print("Загрузка файла input.xlsx...")

input_file <- "input.xlsx"

if (!file.exists(input_file)) {
  stop(paste("Файл", input_file, "не найден в рабочей директории:", getwd()))
}

data <- read.xlsx(input_file, sheet = 1, startRow = 1, rows = 1:40, colNames = TRUE)

print(paste("Успешно загружен файл:", input_file))
print(paste("Загружено строк:", nrow(data), "столбцов:", ncol(data)))

# Проверяем структуру данных
print("Структура данных:")
print(colnames(data))
print(dim(data))

# Если данных меньше 9 столбцов, добавляем недостающие
if(ncol(data) < 9) {
  print(paste("Обнаружено столбцов:", ncol(data), ". Добавляю недостающие..."))
  for(i in (ncol(data)+1):9) {
    data[[paste0("Column", i)]] <- NA
  }
}

# Добавление столбца с номером строки
data <- cbind(data, 1:nrow(data))
colnames(data)[10] <- 'Row_Number'

# Переименование столбцов для удобства
if(ncol(data) >= 9) {
  colnames(data)[1:9] <- c('Intensity_In_1', 'Intensity_Straight_1', 'Intensity_Right_1', 
                           'Green_Time_1', 'Intensity_In_2', 'Intensity_Straight_2', 
                           'Intensity_Right_2', 'Green_Time_2', 'Total_Intensity')
}

print("Столбцы переименованы:")
print(colnames(data))

# ФАЗЗИФИКАЦИЯ
print("Начало фаззификации...")

# Применяем фаззификацию ко всем столбцам
intensity_in_1 <- fuzz_intensity(data, 1)
intensity_straight_1 <- fuzz_intensity(data, 2)
intensity_right_1 <- fuzz_intensity(data, 3)
green_time_1 <- fuzz_green_time(data, 4)
intensity_in_2 <- fuzz_intensity(data, 5)
intensity_straight_2 <- fuzz_intensity(data, 6)
intensity_right_2 <- fuzz_intensity(data, 7)
green_time_2 <- fuzz_green_time(data, 8)
total_intensity <- fuzz_intensity(data, 9)

# Собираем все данные вместе
data_fuzzy <- cbind(
  intensity_in_1[,2:4],
  intensity_straight_1[,2:4],
  intensity_right_1[,2:4],
  green_time_1[,2:4],
  intensity_in_2[,2:4],
  intensity_straight_2[,2:4],
  intensity_right_2[,2:4],
  green_time_2[,2:4],
  total_intensity[,2:4],
  data[,10]
)

colnames(data_fuzzy) <- c(
  'Intensity_In_1_low', 'Intensity_In_1_medium', 'Intensity_In_1_high',
  'Intensity_Straight_1_low', 'Intensity_Straight_1_medium', 'Intensity_Straight_1_high',
  'Intensity_Right_1_low', 'Intensity_Right_1_medium', 'Intensity_Right_1_high',
  'Green_Time_1_low', 'Green_Time_1_medium', 'Green_Time_1_high',
  'Intensity_In_2_low', 'Intensity_In_2_medium', 'Intensity_In_2_high',
  'Intensity_Straight_2_low', 'Intensity_Straight_2_medium', 'Intensity_Straight_2_high',
  'Intensity_Right_2_low', 'Intensity_Right_2_medium', 'Intensity_Right_2_high',
  'Green_Time_2_low', 'Green_Time_2_medium', 'Green_Time_2_high',
  'Total_Intensity_low', 'Total_Intensity_medium', 'Total_Intensity_high',
  'Row_Number'
)

print("Фаззификация завершена!")
print(paste("Фаззифицированные данные:", nrow(data_fuzzy), "строк,", ncol(data_fuzzy), "столбцов"))

# Удаление строк с NA
data_fuzzy <- na.omit(data_fuzzy)
print(paste("После удаления NA:", nrow(data_fuzzy), "строк"))

# ГЕНЕРАЦИЯ КОМПЛЕКСНЫХ ПРАВИЛ ДЛЯ КАЖДОГО МАРШРУТА
print("Генерация комплексных правил для каждого маршрута...")

comprehensive_rules <- c()
human_readable_rules <- c()
priority_rules <- c()

for(i in 1:nrow(data_fuzzy)) {
  cat("Генерация правил для маршрута", i, "\n")
  
  # Метод 1: Полное правило
  rule1 <- generate_comprehensive_rule(data_fuzzy[i, ], data[i, ], i)
  comprehensive_rules <- c(comprehensive_rules, rule1)
  
  # Метод 2: Человеко-понятное правило
  rule2 <- generate_human_readable_rule(data_fuzzy[i, ], data[i, ], i)
  human_readable_rules <- c(human_readable_rules, rule2)
  
  # Метод 3: Приоритетное правило
  rule3 <- generate_priority_rule(data_fuzzy[i, ], data[i, ], i)
  priority_rules <- c(priority_rules, rule3)
}

# НЕЙРОННЫЕ СЕТИ (опционально, для обратной совместимости)
print("Обучение нейросетей...")

regulation_total <- c()
regulations_total <- c()

# Только если данных достаточно
if(nrow(data_fuzzy) >= 10) {
  data_total <- data_fuzzy
  p_total <- vec(data_total, 25, 3)
  data_total <- cbind(data_total[, 1:27], p_total)
  colnames(data_total)[28] <- 'Total_Intensity_Vector'
  
  for(i in 1:2) {
    tryCatch({
      index <- sample(1:nrow(data_total), round(0.10*nrow(data_total)))
      train <- data_total[-index,]
      test <- data_total[index,]
      
      scaled <- scale(data_total)
      train_ <- scaled[-index,]
      test_ <- scaled[index,]
      
      formula_total <- as.formula(paste("Total_Intensity_Vector ~", 
                                        paste(colnames(data_total)[1:24], collapse = " + ")))
      
      rez <- neuralnet(formula_total, data = train_, 
                       algorithm = 'rprop+', 
                       hidden = 2,
                       act.fct = "tanh", 
                       err.fct = "sse", 
                       linear.output = FALSE,
                       stepmax = 1e5)
      
      rule_garson <- generate_rules(rez, "Total_Intensity")
      regulations_total <- c(regulations_total, rule_garson)
      
    }, error = function(e) {
      print(paste("Ошибка в итерации", i, ":", e$message))
    })
  }
} else {
  print("Недостаточно данных для обучения нейросетей")
}

# СОХРАНЕНИЕ РЕЗУЛЬТАТОВ
print("Сохранение результатов...")

wb <- createWorkbook()

# Лист с фаззифицированными данными
addWorksheet(wb, "Фаззифицированные_данные")
writeData(wb, "Фаззифицированные_данные", data_fuzzy)

# Листы с комплексными правилами
addWorksheet(wb, "Полные_правила")
writeData(wb, "Полные_правила", 
          data.frame(Маршрут = 1:length(comprehensive_rules),
                     Правило = comprehensive_rules))

addWorksheet(wb, "Понятные_правила")
writeData(wb, "Понятные_правила",
          data.frame(Маршрут = 1:length(human_readable_rules),
                     Правило = human_readable_rules))

addWorksheet(wb, "Приоритетные_правила")
writeData(wb, "Приоритетные_правила",
          data.frame(Маршрут = 1:length(priority_rules),
                     Правило = priority_rules))

# Лист с нейросетевыми правилами (если есть)
if(length(regulations_total) > 0) {
  addWorksheet(wb, "Нейросетевые_правила")
  writeData(wb, "Нейросетевые_правила",
            data.frame(Номер_правила = 1:length(regulations_total),
                       Правило = regulations_total))
}

# Сохраняем файл
output_file <- "Traffic_Fuzzy_Neural_Results.xlsx"
saveWorkbook(wb, output_file, overwrite = TRUE)

# ВЫВОД СТАТИСТИКИ
print("=== РЕЗУЛЬТАТЫ МОДЕЛИРОВАНИЯ ===")
print(paste("Исходный файл:", input_file))
print(paste("Общее количество наблюдений:", nrow(data_fuzzy)))
print(paste("Количество полных правил:", length(comprehensive_rules)))
print(paste("Количество понятных правил:", length(human_readable_rules)))
print(paste("Количество приоритетных правил:", length(priority_rules)))
if(length(regulations_total) > 0) {
  print(paste("Количество нейросетевых правил:", length(regulations_total)))
}
print(paste("Результаты сохранены в файл:", output_file))

print("Анализ завершен успешно!")

# Пример вывода нескольких правил
if(length(comprehensive_rules) > 0) {
  print("Пример полных правил:")
  for(i in 1:min(3, length(comprehensive_rules))) {
    print(paste(i, ":", comprehensive_rules[i]))
  }
}

if(length(human_readable_rules) > 0) {
  print("Пример понятных правил:")
  for(i in 1:min(3, length(human_readable_rules))) {
    print(paste(i, ":", human_readable_rules[i]))
  }
}
