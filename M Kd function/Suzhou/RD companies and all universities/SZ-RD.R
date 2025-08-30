# Install and load required R packages
if (!requireNamespace("pak", quietly = TRUE)) install.packages("pak")
pak::pak("EricMarcon/dbmss")

if (!requireNamespace("readxl", quietly = TRUE)) install.packages("readxl")
if (!requireNamespace("writexl", quietly = TRUE)) install.packages("writexl")

library(dbmss)
library(readxl)
library(writexl)
library(spatstat.geom)


# =============================
# Load data
# =============================
all_points <- read_excel("D:/2025manchester/ERP/union-RD-SZ-1.xlsx")

# Check required columns
stopifnot(all(c("x_utm", "y_utm", "type") %in% names(all_points)))

# Standardize PointType names
all_points$PointType <- ifelse(all_points$type == "company", "Company", "University")


# =============================
# Create DBMSS-format point pattern
# =============================

# Define window
xrange <- range(all_points$x_utm)
yrange <- range(all_points$y_utm)
W <- owin(xrange = xrange, yrange = yrange)

# Convert to ppp object
pp <- as.ppp(all_points[, c("x_utm", "y_utm")], W)
marks(pp) <- factor(all_points$PointType)

# Convert to DBMSS object
X <- as.wmppp(pp)

# Check summary
summary(X)
table(marks(X))

# =============================
# Compute M function
# =============================
M_result <- Mhat(
  X,
  ReferenceType = "Company",
  NeighborType  = "University"
)

png("M_Function.png", width = 800, height = 600)
plot(M_result, main = "M Function: Company → University")

M_df <- data.frame(r = M_result$r, M = M_result$M)

str(M_result)

getwd()
dev.cur()
dev.off()
library(ggplot2)

ggplot(M_df, aes(x = r, y = M)) +
  geom_line() +
  geom_hline(yintercept = 1, linetype = "dashed", color = "red") +
  coord_cartesian(ylim = c(0.6, 1.1)) +
  labs(title = "M Function: Company → University",
       x = "Distance r (m)", y = "M(r)") +
  theme_minimal()


# =============================
# Compute Kd function
# =============================
Kd_result <- Kdhat(
  X,
  ReferenceType = "Company",
  NeighborType  = "University"
)

png("Kd_Function.png", width = 800, height = 600)
plot(Kd_result, main = "Kd Function: Company → University")
dev.off()

Kd_df <- data.frame(r = Kd_result$r, Kd = Kd_result$Kd)

# =============================
# Kd confidence interval envelope
# =============================
set.seed(123)
envelope_result <- KdEnvelope(
  X,
  ReferenceType = "Company",
  NeighborType  = "University",
)

png("Kd_Envelope.png", width = 800, height = 600)
plot(envelope_result, main = "Kd Envelope: Company → University")
dev.off()

env_df <- data.frame(
  r = envelope_result$r,
  obs = envelope_result$obs,
  lo = envelope_result$lo,
  hi = envelope_result$hi
)

# =============================
# Save results to Excel
# =============================
writexl::write_xlsx(list(
  M_Function = M_df,
  Kd_Function = Kd_df,
  Kd_Envelope = env_df
), "DBMSS_Results.xlsx")

cat("Analysis completed, results saved in current directory:\n",
    "- M_Function.png\n",
    "- Kd_Function.png\n",
    "- Kd_Envelope.png\n",
    "- DBMSS_Results.xlsx\n")

# =============================
# Compute individual M values (local)
# =============================
# Choose a reasonable scale, e.g., 50 km
distance <- 50000  # unit: meters

fvind <- Mhat(
  X,
  r = c(0, distance),
  ReferenceType = "Company",
  NeighborType = "University",
  Individual = TRUE
)

# =============================
# 🌡️ Smooth local M values to generate heatmap
# =============================
p_map <- Smooth(
  X,
  fvind = fvind,
  distance = distance,
  Nbx = 512,
  Nby = 512
)

# =============================
#  Plot heatmap
# =============================
png("M_Local_Heatmap.png", width = 1000, height = 800)
par(mar = rep(0, 4))  # Remove margins
plot(p_map, main = paste("Local M Function Heatmap (", distance/1000, " km)", sep = ""))

# Mark company locations
is_company <- marks(X)$PointType == "Company"
points(
  x = X$x[is_company],
  y = X$y[is_company],
  pch = 20,
  col = "black"
)

# Add contour lines
contour(p_map, nlevels = 5, add = TRUE)
dev.off()

cat("Local M function heatmap saved as: M_Local_Heatmap.png\n")
