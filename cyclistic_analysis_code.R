# Lista todos los archivos .xlsx
files <- list.files(path = "Datos_Cyclistic/", pattern = "*.xlsx", full.names = TRUE)

read_clean <- function(file) {
  read_excel(file) %>%
    clean_names() %>%
    mutate(
      end_station_id = as.character(end_station_id),
      start_station_id = as.character(start_station_id)
    )
}

cyclistic_data <- files %>%
  map_df(read_clean)

cyclistic_data <- cyclistic_data %>%
  mutate(
    started_at = ymd_hms(started_at),
    ended_at = ymd_hms(ended_at),
    ride_length = as.numeric(difftime(ended_at, started_at, units = "secs")),
    day_of_week = wday(started_at, label = TRUE, abbr = FALSE), # i.e. Monday
    month = month(started_at, label = TRUE, abbr = FALSE)
  ) %>%
  filter(ride_length > 0 & ride_length < 86400) # Filtering outliers

cyclistic_data %>%
  group_by(member_casual) %>%
  summarise(avg_duration = mean(ride_length) / 60) # For minutes

cyclistic_data %>%
  group_by(member_casual) %>%
  summarise(
    total_rides = n(),
    median_duration = median(ride_length),
    max_duration = max(ride_length),
    min_duration = min(ride_length)
  )

cyclistic_data %>%
  count(month, member_casual) %>%
  arrange(match(month, month.name)) %>%
  print(n = 24)

cyclistic_data %>%
  count(day_of_week, member_casual)

ggplot(cyclistic_data, aes(x = day_of_week, fill = member_casual)) +
  geom_bar(position = "dodge") +
  labs(title = "Daily rides", x = "Day of week", y = "Rides")

cyclistic_data %>%
  group_by(member_casual) %>%
  summarise(avg_minutes = mean(ride_length) / 60) %>%
  ggplot(aes(x = member_casual, y = avg_minutes, fill = member_casual)) +
  geom_col() +
  labs(title = "AVG duration by user", y = "Mins")

graphic1 <- ggplot(cyclistic_data, aes(x = day_of_week, fill = member_casual)) +
  geom_bar(position = "dodge") +
  labs(title = "Daily rides", x = "Day of week", y = "Rides")

graphic2 <- cyclistic_data %>%
  group_by(member_casual) %>%
  summarise(avg_minutes = mean(ride_length) / 60) %>%
  ggplot(aes(x = member_casual, y = avg_minutes, fill = member_casual)) +
  geom_col() +
  labs(title = "AVG duration by user", y = "Mins")

ggsave("rides_by_day.png", plot = graphic1, width = 8, height = 5, dpi = 300)

ggsave("avg_duration_by_user.png", plot = graphic2, width = 8, height = 5, dpi = 300)

duration_summary <- cyclistic_data %>%
  group_by(member_casual) %>%
  summarise(promedio_min = mean(ride_length) / 60)

statistics <- cyclistic_data %>%
  group_by(member_casual) %>%
  summarise(
    total_rides = n(),
    median_duration = median(ride_length),
    max_duration = max(ride_length),
    min_duration = min(ride_length)
  )

rides_by_month <- cyclistic_data %>%
  count(month, member_casual)

rides_by_day <- cyclistic_data %>%
  count(day_of_week, member_casual)

xlsx <- createWorkbook()

addWorksheet(xlsx, "AVG Duration")
writeData(xlsx, "AVG Duration", duration_summary)

addWorksheet(xlsx, "Rides by month")
writeData(xlsx, "Rides by month", rides_by_month)

addWorksheet(xlsx, "Rides by day")
writeData(xlsx, "Rides by day", rides_by_day)

saveWorkbook(xlsx, "cyclistic_summary.xlsx", overwrite = TRUE)

xlsx2 <- createWorkbook()

addWorksheet(xlsx2, "Statistics")
writeData(xlsx2, "Statistics", statistics)

saveWorkbook(xlsx2, "statistics.xlsx", overwrite = TRUE)

write.xlsx(cyclistic_data, file = "cyclistic_clean_database.xlsx", overwrite = TRUE)