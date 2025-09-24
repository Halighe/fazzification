library(XLConnect)
library(neuralnet)
library(NeuralNetTools)

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
  View(arr)
}

# Функция фазификации для 3 переменных
fuzz <- function(data, numcol){
  fuzz_means <- array(0, dim = length(data[,numcol]))
  fuzz_max <- array(0, dim = length(data[,numcol]))
  for(f in 1:length(data[,numcol])){
    if(data[f,numcol] > 1 & data[f,numcol] <= 2){
      fuzz_means[f] <- 1 # 2
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] > 2 & data[f,numcol] <= 3){
      fuzz_max[f] <- 1 # 3
      data[f,numcol] <- 0
    }
    else {#
      data[f,numcol] <- 1
    }
  }
  c <- cbind(data[, numcol], fuzz_means, fuzz_max)
  return(c)
}

# Функция фазификации для 4 переменных
fuzz1 <- function(data, numcol){
  fuzz_Belmeans <- array(0, dim = length(data[,numcol]))
  fuzz_Abomeans <- array(0, dim = length(data[,numcol]))
  fuzz_max <- array(0, dim = length(data[,numcol]))
  for(f in 1:length(data[,numcol])){
    if(data[f,numcol] > 1 & data[f,numcol] <= 2){
      fuzz_Belmeans[f] <- 1 # 2
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] > 2 & data[f,numcol] <= 3){
      fuzz_Abomeans[f] <- 1 # 3
      data[f,numcol] <- 0
    }
    else if(data[f,numcol] > 3 & data[f,numcol] <= 4) {#
      fuzz_max[f] <- 1 # 3
      data[f,numcol] <- 0
    }
    else {#
      data[f,numcol] <- 1
    }
  }
  c <- cbind(data[, numcol], fuzz_Belmeans, fuzz_Abomeans, fuzz_max)
  return(c)
}

# Функция преобразовния в вектор
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

wd <- loadWorkbook(file.choose(), create = T)
data <- readWorksheet(wd, sheet = "Лист1", startRow = 1, endRow = 39, 
                     startCol = 0, endCol = 9)

# Добавление столбца с номером строки и удаление выбросов 
data <- cbind(data, 1:length(data[, 1]))
colnames(data)[10] <- 'Col-vo'

# Переход к нечеткому множетсву для Capacity
cap_min <- min(data[, 9]) - 1000
cap_max <- max(data[, 9]) / 3
data_cap <- data[data$Capacity <= cap_max, c(9,10)]
cap_mean <- mean(data_cap[, 1])
fuzzy_cap <- f(cap_min, cap_max, cap_mean, data_cap, 0)

data_cap1 <- data[data$Capacity > cap_max & data$Capacity <= cap_max * 2, c(9,10)]
cap_mean1 <- mean(data_cap1[, 1])
fuzzy_cap1 <- f(cap_max, cap_max * 2, cap_mean1, data_cap1, 1)

data_cap2 <- data[data$Capacity > cap_max * 2 & data$Capacity <= cap_max * 3, c(9,10)]
cap_mean2 <- mean(data_cap2[, 1])
fuzzy_cap2 <- f(cap_max*2, cap_max * 3, cap_mean2, data_cap2, 2)

x <- rbind(fuzzy_cap, fuzzy_cap1, fuzzy_cap2)
data[, 9] <- x[order(x[,2]), 1]

# Переход к нечеткому множетсву для Grade
mins <- min(data[, 2]) - 1
maxs <- max(data[, 2]) / 3
data_ <- data[data$Grade <= maxs, c(2,10)]
means <- mean(data_[, 1])
fuzzy <- f(mins, maxs, means, data_, 0)

data1 <- data[data$Grade > maxs & data$Grade <= maxs * 2, c(2,10)]
means1 <- mean(data1[, 1])
fuzzy1 <- f(maxs, maxs * 2, means1, data1, 1)

data2 <- data[data$Grade > maxs * 2 & data$Grade <= maxs * 3, c(2,10)]
means2 <- mean(data2[, 1])
fuzzy2 <- f(maxs*2, maxs * 3, means2, data2, 2)

x <- rbind(fuzzy, fuzzy1, fuzzy2)
data[, 2] <- x[order(x[,2]), 1]

# Переход к нечеткому множетсву для Trucks
mins <- min(data[, 3]) - 3
maxs <- max(data[, 3]) / 3
data_ <- data[data$Trucks <= maxs, c(3,10)]
means <- mean(data_[, 1])
fuzzy <- f(mins, maxs, means, data_, 0)

data1 <- data[data$Trucks > maxs & data$Trucks <= maxs * 2, c(3,10)]
means1 <- mean(data1[, 1])
fuzzy1 <- f(maxs, maxs * 2, means1, data1, 1)

data2 <- data[data$Trucks > maxs * 2 & data$Trucks <= maxs * 3, c(3,10)]
means2 <- mean(data2[, 1])
fuzzy2 <- f(maxs*2, maxs * 3, means2, data2, 2)

x <- rbind(fuzzy, fuzzy1, fuzzy2)
data[, 3] <- x[order(x[,2]), 1]

# Переход к нечеткому множетсву для NumLanes
#mins <- min(data[, 5])
#maxs <- max(data[, 5]) / 3
#data_ <- data[data$NumLanes <= maxs, c(5,10)]
#means <- mean(data_[, 1])
#fuzzy <- f(mins, maxs, means, data_, 0)
#
#data1 <- data[data$NumLanes > maxs & data$NumLanes <= maxs * 2, c(5,10)]
#means1 <- mean(data1[, 1])
#fuzzy1 <- f(maxs, maxs * 2, means1, data1, 1)
#
#data2 <- data[data$NumLanes > maxs * 2 & data$NumLanes <= maxs * 3, c(5,10)]
#means2 <- mean(data2[, 1])
#fuzzy2 <- f(maxs*2, maxs * 3, means2, data2, 2)
#
#x <- rbind(fuzzy, fuzzy1, fuzzy2)
#data[, 5] <- x[order(x[,2]), 1]

# Переход к нечеткому множетсву для Width
mins <- min(data[, 6])
maxs <- max(data[, 6]) / 3
data_ <- data[data$Width <= maxs, c(6,10)]
means <- mean(data_[, 1])
fuzzy <- f(mins, maxs, means, data_, 0)

data1 <- data[data$Width > maxs & data$Width <= maxs * 2, c(6,10)]
means1 <- mean(data1[, 1])
fuzzy1 <- f(maxs, maxs * 2, means1, data1, 1)

data2 <- data[data$Width > maxs * 2 & data$Width <= maxs * 3, c(6,10)]
means2 <- mean(data2[, 1])
fuzzy2 <- f(maxs*2, maxs * 3, means2, data2, 2)

x <- rbind(fuzzy, fuzzy1, fuzzy2)
data[, 6] <- x[order(x[,2]), 1]

# Удаление столбца Work
data_fin <- cbind(data[, 1:3], data[, 5:9])

data <- data_fin

# Фазификация для Location, Grade, Trucks, Width, Capacity
a <- fuzz(data, 1)
colnames(a)[1] <- 'Location_rural'
colnames(a)[2] <- 'Location_urban'

g <- fuzz(data, 2)
colnames(g)[1] <- 'Grade_min'
colnames(g)[2] <- 'Grade_means'
colnames(g)[3] <- 'Grade_max'

t <- fuzz(data, 3)
colnames(t)[1] <- 'Trucks_min'
colnames(t)[2] <- 'Trucks_means'
colnames(t)[3] <- 'Trucks_max'

z <- fuzz1(data, 4)
colnames(z)[1] <- 'NumLanes_min'
colnames(z)[2] <- 'NumLanes_Belmeans'
colnames(z)[3] <- 'NumLanes_Abomeans'
colnames(z)[4] <- 'NumLanes_max'

w <- fuzz(data, 5)
colnames(w)[1] <- 'Width_min'
colnames(w)[2] <- 'Width_means'
colnames(w)[3] <- 'Width_max'

v <- fuzz(data, 6)
colnames(v)[1] <- 'Reduced_yes'
colnames(v)[2] <- 'Reduced_no'

b <- fuzz(data, 7)
colnames(b)[1] <- 'SpeedLimit_min'
colnames(b)[2] <- 'SpeedLimit_max'

c <- fuzz(data, 8)
colnames(c)[1] <- 'Capacity_min'
colnames(c)[2] <- 'Capacity_means'
colnames(c)[3] <- 'Capacity_max'

data <- cbind(a[,1:2], g, t, z[,1:4], w, v[,1:2], b[,1:2], c)
#colnames(data)[9] <- 'NumLanes'
#View(data)

# Удаление дубликатов
kol <- 0
koef <- 0
j <- 0
len <- length(data[,1])
for(h in 1:(len-1)){
  for(k in (h+1):len){
    for(i in 1:length(data[1,])){
      if(Mod(data[h-j, i] - data[k-j, i]) <= koef) {kol <- kol + 1}
    }
    if(kol == 22) {
      data <- data[-(h-j),]
      j <- j + 1
      kol <- 0
      break
    }
    kol <- 0
  }
}
data_past <- data


# Короткий вектор Trucks
data <- data_past
p <- vec(data, 5, 3)
data <- cbind(data[, 1:22], p)
colnames(data)[23] <- 'Trucks'

# Нейросеть и формирование правил с выходом Trucks
regulation <- c()
regulations <- c()
#prov <- c()
#while(is.element('Trucks=min', prov) == FALSE || is.element('Trucks=means', prov) == FALSE || is.element('Trucks=max', prov) == FALSE){
for(i in 1:10){
  # Скалирование, разбиение на тестовую и обучающую выборки
  index <- sample(1:nrow(data), round(0.15*nrow(data)))
  train <- data[-index,]
  test <- data[index,]
  
  scaled <- (scale(data))
  train_ <- scaled[-index,]
  test_ <- scaled[index,]
  
  # Нейросеть для Trucks с малым вектором
  rez <- neuralnet(Trucks ~ Location_rural + Location_urban + Grade_min + Grade_means + Grade_max + Capacity_min + Capacity_means + Capacity_max + NumLanes_min + NumLanes_Belmeans + NumLanes_Abomeans + NumLanes_max + Width_min + Width_means + Width_max + Reduced_yes + Reduced_no + Trucks_min + Trucks_means + Trucks_max + SpeedLimit_min + SpeedLimit_max, data = train_, algorithm='rprop+',hidden=5, act.fct="tanh", err.fct = "sse", linear.output = F)
  un <- train_
  un[,23] <- as.numeric(rez$net.result[[1]][,1])
  train_ <- scale(train)
  scaleList <- list(scale = attr(train_, "scaled:scale"), center = attr(train_, "scaled:center"))
  un[,23]*scaleList$scale["Trucks"] + scaleList$center["Trucks"]
  comp1 <- data.frame(ish=train[,23], fit=un[,23]*scaleList$scale["Trucks"] + scaleList$center["Trucks"])
  #plot(rez)
  
  # Формирование правил с помощью альтернативного метода
  weght_hide <- rez$weights[[1]][[2]]
  weght_hide <- weght_hide[-1,]
  weght_hide <- which.max(weght_hide == max(weght_hide))
  
  alt_method <- rez$weights[[1]][[1]]
  alt_method <- alt_method[-1,]
  alt_method <- alt_method[,weght_hide]
  
  vx <- rez$model.list$variables
  sr <- array(0, dim = 8)
  
  for(j in 1:8){
    ind <- which.max(alt_method == max(alt_method))
    punkt <- vx[ind] 
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, vx)
    vx <- vx[-c]
    alt_method <- alt_method[-c]
    
    num_s <- gregexpr("_", punkt)[[1]][1]
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    #if(j < 8){x <- length(pr_sort[,1])}
    
  }
  zi <- sr[grep("Trucks", sr)]
  sr <- sr[-grep("Trucks", sr)]
  pr <- "если:"
  for(i in 1:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", zi)
  #View(pr)
  regulation <- c(regulation, pr)
  
  
  # Формирование правил с помощью метода Гарсона
  w <- garson(rez)
  #plot(w)
  pr_sort <- table(w$data[order(w$data[,1]),])
  #pr_sort <- pr_sort[-match("NumLanes", colnames(pr_sort)), -match("NumLanes", colnames(pr_sort))]
  sr <- array(0, dim = 8)
  x <- max(grep("Trucks", colnames(pr_sort)))
  j <- 1
  for(j in 1:8){
    punkt <- colnames(pr_sort)[x]
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, colnames(pr_sort))
    pr_sort <- pr_sort[-c, -c]
    
    num_s <- gregexpr("_", punkt)[[1]][1] 
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    if(j < 8){x <- length(pr_sort[,1])}
  }
  #View(sr)
  
  pr <- "если:"
  for(i in 2:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", sr[1])
  #View(pr)
  #prov <- c(prov, sr[1]) 
  regulations <- c(regulations, pr)
}  
writeWorksheetToFile("Test.xlsx", data = cbind(regulation, regulations), sheet = "Лист1")


# Короткий вектор SpeedLimit
data <- data_past
p <- vec(data, 17, 2)
data <- cbind(data[, 1:22], p)
colnames(data)[23] <- 'SpeedLimit'

# Нейросеть и формирование правил с выходом SpeedLimit
regulation1 <- c()
regulations1 <- c()
#prov <- c()
#while(is.element('SpeedLimit=min', prov) == FALSE || is.element('SpeedLimit=max', prov) == FALSE){
for(i in 1:10){
  # Скалирование, разбиение на тестовую и обучающую выборки
  index <- sample(1:nrow(data), round(0.15*nrow(data)))
  train <- data[-index,]
  test <- data[index,]
  
  scaled <- (scale(data))
  train_ <- scaled[-index,]
  test_ <- scaled[index,]
  
  # Нейросеть для SpeedLimit с малым вектором
  rez <- neuralnet(SpeedLimit ~ Location_rural + Location_urban + Grade_min + Grade_means + Grade_max + Capacity_min + Capacity_means + Capacity_max + NumLanes_min + NumLanes_Belmeans + NumLanes_Abomeans + NumLanes_max + Width_min + Width_means + Width_max + Reduced_yes + Reduced_no + SpeedLimit_min + SpeedLimit_max + Trucks_min + Trucks_means + Trucks_max, data = train_, algorithm='rprop+',hidden=5, act.fct="tanh", err.fct = "sse", linear.output = F)
  un <- train_
  un[,23] <- as.numeric(rez$net.result[[1]][,1])
  train_ <- scale(train)
  scaleList <- list(scale = attr(train_, "scaled:scale"), center = attr(train_, "scaled:center"))
  un[,23]*scaleList$scale["SpeedLimit"] + scaleList$center["SpeedLimit"]
  comp1 <- data.frame(ish=train[,23], fit=un[,23]*scaleList$scale["SpeedLimit"] + scaleList$center["SpeedLimit"])
  #plot(rez)
  
  # Формирование правил с помощью альтернативного метода
  weght_hide <- rez$weights[[1]][[2]]
  weght_hide <- weght_hide[-1,]
  weght_hide <- which.max(weght_hide == max(weght_hide))
  
  alt_method <- rez$weights[[1]][[1]]
  alt_method <- alt_method[-1,]
  alt_method <- alt_method[,weght_hide]
  
  vx <- rez$model.list$variables
  sr <- array(0, dim = 8)
  
  for(j in 1:8){
    ind <- which.max(alt_method == max(alt_method))
    punkt <- vx[ind] 
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, vx)
    vx <- vx[-c]
    alt_method <- alt_method[-c]
    
    num_s <- gregexpr("_", punkt)[[1]][1]
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    #if(j < 8){x <- length(pr_sort[,1])}
    
  }
  zi <- sr[grep("SpeedLimit", sr)]
  sr <- sr[-grep("SpeedLimit", sr)]
  pr <- "если:"
  for(i in 1:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", zi)
  #View(pr)
  regulation1 <- c(regulation1, pr)
  
  
  # Формирование правил с помощью метода Гарсона
  w <- garson(rez)
  #plot(w)
  pr_sort <- table(w$data[order(w$data[,1]),])
  #pr_sort <- pr_sort[-match("NumLanes", colnames(pr_sort)), -match("NumLanes", colnames(pr_sort))]
  sr <- array(0, dim = 8)
  x <- max(grep("SpeedLimit", colnames(pr_sort)))
  for(j in 1:8){
    punkt <- colnames(pr_sort)[x]
    
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, colnames(pr_sort))
    pr_sort <- pr_sort[-c, -c]
    
    num_s <- gregexpr("_", punkt)[[1]][1] 
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    if(j < 8){x <- length(pr_sort[,1])}
  }
  #View(sr)
  
  pr <- "если:"
  for(i in 2:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", sr[1])
  #View(pr)
  #prov <- c(prov, sr[1]) 
  regulations1 <- c(regulations1, pr)
}  
writeWorksheetToFile("Test1.xlsx", data = cbind(regulation1, regulations1), sheet = "Лист1")


# Короткий вектор NumLanes
data <- data_past
p <- vec(data, 8, 4)
data <- cbind(data[, 1:22], p)
colnames(data)[23] <- 'NumLanes'

# Нейросеть и формирование правил с выходом NumLanes
regulation2 <- c()
regulations2 <- c()
#prov <- c()
data <- data[-1,-9]
data <- data[-c(28,29),-11]
data <- data[,-18]
#while(is.element('SpeedLimit=min', prov) == FALSE || is.element('SpeedLimit=max', prov) == FALSE){
for(i in 1:20){
  # Скалирование, разбиение на тестовую и обучающую выборки
  index <- sample(1:nrow(data), round(0.15*nrow(data)))
  train <- data[-index,]
  test <- data[index,]
  
  scaled <- (scale(data))
  train_ <- scaled[-index,]
  test_ <- scaled[index,]
  
  # Нейросеть для NumLanes с малым вектором
  rez <- neuralnet(NumLanes ~ Location_rural + Location_urban + Grade_min + Grade_means + Grade_max + Capacity_means + Capacity_max + NumLanes_Belmeans + NumLanes_Abomeans + Width_min + Width_means + Width_max + Reduced_yes + Reduced_no + SpeedLimit_min + SpeedLimit_max + Trucks_min + Trucks_means + Trucks_max, data = train_, algorithm='rprop+',hidden=5, act.fct="tanh", err.fct = "sse", linear.output = F)
  un <- train_
  un[,20] <- as.numeric(rez$net.result[[1]][,1])
  train_ <- scale(train)
  scaleList <- list(scale = attr(train_, "scaled:scale"), center = attr(train_, "scaled:center"))
  un[,20]*scaleList$scale["NumLanes"] + scaleList$center["NumLanes"]
  comp1 <- data.frame(ish=train[,20], fit=un[,20]*scaleList$scale["NumLanes"] + scaleList$center["NumLanes"])
  #plot(rez)
  
  # Формирование правил с помощью альтернативного метода
  weght_hide <- rez$weights[[1]][[2]]
  weght_hide <- weght_hide[-1,]
  weght_hide <- which.max(weght_hide == max(weght_hide))
  
  alt_method <- rez$weights[[1]][[1]]
  alt_method <- alt_method[-1,]
  alt_method <- alt_method[,weght_hide]
  
  vx <- rez$model.list$variables
  sr <- array(0, dim = 8)
  
  for(j in 1:8){
    ind <- which.max(alt_method == max(alt_method))
    punkt <- vx[ind] 
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, vx)
    vx <- vx[-c]
    alt_method <- alt_method[-c]
    
    num_s <- gregexpr("_", punkt)[[1]][1]
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    #if(j < 8){x <- length(pr_sort[,1])}
    
  }
  zi <- sr[grep("NumLanes", sr)]
  sr <- sr[-grep("NumLanes", sr)]
  pr <- "если:"
  for(i in 1:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", zi)
  #View(pr)
  regulation2 <- c(regulation2, pr)
  
  
  # Формирование правил с помощью метода Гарсона
  w <- garson(rez)
  #plot(w)
  pr_sort <- table(w$data[order(w$data[,1]),])
  #pr_sort <- pr_sort[-match("NumLanes", colnames(pr_sort)), -match("NumLanes", colnames(pr_sort))]
  sr <- array(0, dim = 8)
  x <- max(grep("NumLanes", colnames(pr_sort)))
  for(j in 1:8){
    punkt <- colnames(pr_sort)[x]
    
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, colnames(pr_sort))
    pr_sort <- pr_sort[-c, -c]
    
    num_s <- gregexpr("_", punkt)[[1]][1] 
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    if(j < 8){x <- length(pr_sort[,1])}
  }
  #View(sr)
  
  pr <- "если:"
  for(i in 2:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", sr[1])
  #View(pr)
  #prov <- c(prov, sr[1]) 
  regulations2 <- c(regulations2, pr)
}
writeWorksheetToFile("Test2.xlsx", data = cbind(regulation2, regulations2), sheet = "Лист1")



#writeWorksheetToFile("Test.xlsx", data = c(regulations, regulations1, regulations2), sheet = "Лист1")


# Удаление строк где нет полосы для разгона
data <- data_past
j <- c()
for(g in 1:length(data_past[,1])){
  if(data_past[g,16] == 0){
    j <- c(j, g)
  }
}
data <- cbind(data_past[-j, 1:10], data[-j, 13:14], data_past[-j, 18:21])


# Короткий вектор NumLanes
#data <- data_past
p <- vec(data, 8, 2)
data <- cbind(data[, 1:16], p)
colnames(data)[17] <- 'NumLanes'

# Нейросеть и формирование правил с выходом NumLanes
regulations2 <- c()
prov <- c()
#while(is.element('NumLanes=min', prov) == FALSE || is.element('NumLanes=Belmeans', prov) == FALSE){
for(i in 1:10){
  # Скалирование, разбиение на тестовую и обучающую выборки
  index <- sample(1:nrow(data), round(0.15*nrow(data)))
  train <- data[-index,]
  test <- data[index,]
  
  scaled <- (scale(data))
  train_ <- scaled[-index,]
  test_ <- scaled[index,]
  
  # Нейросеть для NumLanes с малым вектором
  rez <- neuralnet(NumLanes ~ Location_rural + Location_urban + Grade_min + Grade_means + Grade_max + Capacity_min + Capacity_means + NumLanes_min + NumLanes_Belmeans + Width_min + Width_means + Trucks_min + Trucks_means + Trucks_max + SpeedLimit_min + SpeedLimit_max, data = train_, algorithm='rprop+',hidden=5, act.fct="tanh", err.fct = "sse", linear.output = F)
  un <- train_
  un[,17] <- as.numeric(rez$net.result[[1]][,1])
  train_ <- scale(train)
  scaleList <- list(scale = attr(train_, "scaled:scale"), center = attr(train_, "scaled:center"))
  un[,17]*scaleList$scale["NumLanes"] + scaleList$center["NumLanes"]
  comp1 <- data.frame(ish=train[,17], fit=un[,17]*scaleList$scale["NumLanes"] + scaleList$center["NumLanes"])
  #plot(rez)
  w <- garson(rez)
  #plot(w)
  
  # Формирование правил с помощью метода Гарсона
  pr_sort <- table(w$data[order(w$data[,1]),])
  #pr_sort <- pr_sort[-match("NumLanes", colnames(pr_sort)), -match("NumLanes", colnames(pr_sort))]
  sr <- array(0, dim = 7)
  x <- max(grep("NumLanes", colnames(pr_sort)))
  for(j in 1:7){
    punkt <- colnames(pr_sort)[x]
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, colnames(pr_sort))
    pr_sort <- pr_sort[-c, -c]
    
    num_s <- gregexpr("_", punkt)[[1]][1] 
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    if(j < 7){x <- length(pr_sort[,1])}
  }
  #View(sr)
  
  pr <- "если:"
  for(i in 2:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", sr[1])
  #View(pr)
  prov <- c(prov, sr[1]) 
  regulations2 <- c(regulations2, pr)
} 
writeWorksheetToFile("Test.xlsx", data = c(regulations, regulations1, regulations2), sheet = "Лист1")


# Короткий вектор Reduced
#data <- data_past
p <- vec(data, 15, 2)
data <- cbind(data[, 1:22], p)
colnames(data)[23] <- 'Reduced'

# Нейросеть и формирование правил с выходом Reduced
regulations2 <- c()
prov <- c()
#while(is.element('Reduced=yes', prov) == FALSE || is.element('Reduced=no', prov) == FALSE){
for(i in 1:10){
  # Скалирование, разбиение на тестовую и обучающую выборки
  index <- sample(1:nrow(data), round(0.15*nrow(data)))
  train <- data[-index,]
  test <- data[index,]
  
  scaled <- (scale(data))
  train_ <- scaled[-index,]
  test_ <- scaled[index,]
  
  # Нейросеть для Reduced с малым вектором
  rez <- neuralnet(Reduced ~ Location_rural + Location_urban + Grade_min + Grade_means + Grade_max + Capacity_min + Capacity_means + Capacity_max + NumLanes_min + NumLanes_Belmeans + NumLanes_Abomeans + NumLanes_max + Width_min + Width_means + Width_max + Reduced_yes + Reduced_no + Trucks_min + Trucks_means + Trucks_max + SpeedLimit_min + SpeedLimit_max, data = train_, algorithm='rprop+',hidden=5, act.fct="tanh", err.fct = "sse", linear.output = F)
  un <- train_
  un[,23] <- as.numeric(rez$net.result[[1]][,1])
  train_ <- scale(train)
  scaleList <- list(scale = attr(train_, "scaled:scale"), center = attr(train_, "scaled:center"))
  un[,23]*scaleList$scale["Reduced"] + scaleList$center["Reduced"]
  comp1 <- data.frame(ish=train[,23], fit=un[,23]*scaleList$scale["Reduced"] + scaleList$center["Reduced"])
  #plot(rez)
  w <- garson(rez)
  #plot(w)
  
  # Формирование правил с помощью метода Гарсона
  pr_sort <- table(w$data[order(w$data[,1]),])
  #pr_sort <- pr_sort[-match("NumLanes", colnames(pr_sort)), -match("NumLanes", colnames(pr_sort))]
  sr <- array(0, dim = 8)
  x <- max(grep("Reduced", colnames(pr_sort)))
  for(j in 1:8){
    punkt <- colnames(pr_sort)[x]
    symblos <- substr(punkt, start = 1, stop = 5)
    c <- grep(symblos, colnames(pr_sort))
    pr_sort <- pr_sort[-c, -c]
    
    num_s <- gregexpr("_", punkt)[[1]][1] 
    substr(punkt, start = num_s, stop = num_s) <- "="
    sr[j] <- punkt
    if(j < 8){x <- length(pr_sort[,1])}
  }
  #View(sr)
  
  pr <- "если:"
  for(i in 2:length(sr)){
    pr <- paste(pr, sr[i]) 
  }
  pr <- paste(pr, "то:", sr[1])
  #View(pr)
  prov <- c(prov, sr[1]) 
  regulations2 <- c(regulations2, pr)
} 


#RMSE(comp[,2], comp[,1], length(comp[,1]))
#RMSE(comp1[,2], comp1[,1], length(comp1[,1]))
#RMSE(comp2[,2], comp2[,1], length(comp2[,1]))