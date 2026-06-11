import pandas as pd
import numpy as np
import os
import matplotlib.pyplot as plt
import seaborn as sns
from scipy.stats import pearsonr
import sys

# Ensure correct console output encoding
sys.stdout.reconfigure(encoding='utf-8')

# Input paths
satisfaction_file = r"c:\Users\user\Desktop\AI_test\data\삶의_만족도_시도__20260606195059.xlsx"
suicide_file = r"c:\Users\user\Desktop\AI_test\data\인구십만명당_자살률_시도_시_군_구__20260606194913.xlsx"

# Output directory (Artifacts folder)
artifact_dir = "."
os.makedirs(artifact_dir, exist_ok=True)

# Set up matplotlib for Korean font support
plt.rcParams['font.family'] = 'Malgun Gothic'
plt.rcParams['axes.unicode_minus'] = False
sns.set_theme(style="whitegrid", font='Malgun Gothic')

# ==========================================
# 1. DATA PREPROCESSING
# ==========================================
print("Loading and cleaning data...")

# A. Satisfaction Data
df_sat_raw = pd.read_excel(satisfaction_file, sheet_name='데이터')
cols_sat_raw = df_sat_raw.columns.tolist()

# Standardize year names
years_sat = []
for col in cols_sat_raw:
    if col.startswith('행정구역별') or col.startswith('특성별'):
        years_sat.append(col)
    else:
        years_sat.append(col.split('.')[0])

df_sat_data = df_sat_raw.iloc[1:].copy()
df_sat_data['행정구역별(1)'] = df_sat_data['행정구역별(1)'].ffill()
df_sat_data['특성별(1)'] = df_sat_data['특성별(1)'].ffill()

sat_rows = []
for idx, row in df_sat_data.iterrows():
    region = row.iloc[0]
    feat1 = row.iloc[1]
    feat2 = row.iloc[2]
    
    if region == '행정구역별(1)' or pd.isna(region):
        continue
        
    if feat1 == '전체' and feat2 == '계':
        gender = '계'
    elif feat1 == '성별' and feat2 == '남자':
        gender = '남자'
    elif feat1 == '성별' and feat2 == '여자':
        gender = '여자'
    else:
        continue
        
    for i in range(3, len(cols_sat_raw), 6):
        year = years_sat[i]
        sat_rows.append({
            '행정구역': region,
            '성별': gender,
            '연도': int(year),
            '만족도_계': row.iloc[i],
            '매우_만족': row.iloc[i+1],
            '약간_만족': row.iloc[i+2],
            '보통': row.iloc[i+3],
            '약간_불만족': row.iloc[i+4],
            '매우_불만족': row.iloc[i+5]
        })

df_sat_clean = pd.DataFrame(sat_rows)

# B. Suicide Data
df_sui_raw = pd.read_excel(suicide_file, sheet_name='데이터')
cols_sui_raw = df_sui_raw.columns.tolist()

years_sui = []
for col in cols_sui_raw:
    if col.startswith('행정구역별'):
        years_sui.append(col)
    else:
        years_sui.append(col.split('.')[0])

df_sui_data = df_sui_raw.iloc[1:].copy()

sui_rows = []
for idx, row in df_sui_data.iterrows():
    region = row.iloc[0]
    if region == '행정구역별(1)' or pd.isna(region):
        continue
    
    # Standardize region names
    if region == '전라북도':
        region = '전북특별자치도'
    elif region == '제주도':
        region = '제주특별자치도'
        
    for i in range(1, len(cols_sui_raw), 3):
        year = years_sui[i]
        sui_rows.append({
            '행정구역': region,
            '연도': int(year),
            '계': row.iloc[i],
            '남자': row.iloc[i+1],
            '여자': row.iloc[i+2]
        })

df_sui_clean_raw = pd.DataFrame(sui_rows)

# Melt Suicide Data
sui_melted = []
for idx, row in df_sui_clean_raw.iterrows():
    region = row['행정구역']
    year = row['연도']
    for g in ['계', '남자', '여자']:
        sui_melted.append({
            '행정구역': region,
            '연도': year,
            '성별': g,
            '자살률': row[g]
        })
df_sui_clean = pd.DataFrame(sui_melted)

# Standardize Satisfaction regions as well
df_sat_clean['행정구역'] = df_sat_clean['행정구역'].replace({
    '전라북도': '전북특별자치도',
    '제주도': '제주특별자치도'
})

# C. Merge and clean final types
df_merged = pd.merge(df_sat_clean, df_sui_clean, on=['행정구역', '연도', '성별'], how='inner')

# Force numeric conversions
numeric_cols = ['만족도_계', '매우_만족', '약간_만족', '보통', '약간_불만족', '매우_불만족', '자살률']
for col in numeric_cols:
    df_merged[col] = pd.to_numeric(df_merged[col], errors='coerce')

# Drop any row with missing value in crucial variables
df_merged = df_merged.dropna(subset=['매우_만족', '약간_만족', '자살률'])

# Add helper variables
df_merged['만족율'] = df_merged['매우_만족'] + df_merged['약간_만족']
df_merged['불만족율'] = df_merged['약간_불만족'] + df_merged['매우_불만족']

# Save cleaned data to CSV
cleaned_csv_path = os.path.join(artifact_dir, "cleaned_data.csv")
df_merged.to_csv(cleaned_csv_path, index=False, encoding='utf-8-sig')
print(f"Saved cleaned data to {cleaned_csv_path}")

# Separate National and Regional data
df_national = df_merged[df_merged['행정구역'] == '전국'].copy()
df_regional = df_merged[df_merged['행정구역'] != '전국'].copy()

# ==========================================
# 2. STATISTICAL CORRELATION ANALYSIS
# ==========================================
print("\nPerforming correlation analysis on regional data (excluding '전국')...")

satisfaction_variables = ['매우_만족', '약간_만족', '만족율', '보통', '약간_불만족', '매우_불만족', '불만족율']

# A. Overall correlation
overall_corr = {}
print("\n[Overall Correlation (All years, regions, gender groups combined)]")
print(f"{'변수':<15} | {'상관계수 (r)':<12} | {'p-value':<12} | {'유의성':<8}")
print("-" * 55)
for var in satisfaction_variables:
    r, p = pearsonr(df_regional[var], df_regional['자살률'])
    sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "n.s."
    overall_corr[var] = {'r': r, 'p': p, 'sig': sig}
    print(f"{var:<15} | {r:>12.4f} | {p:>12.4e} | {sig:<8}")

# B. Correlation by Gender
gender_corr = {}
print("\n[Correlation by Gender]")
for gender in ['계', '남자', '여자']:
    print(f"\n성별: {gender}")
    print(f"{'변수':<15} | {'상관계수 (r)':<12} | {'p-value':<12} | {'유의성':<8}")
    print("-" * 55)
    gender_corr[gender] = {}
    df_g = df_regional[df_regional['성별'] == gender]
    for var in satisfaction_variables:
        r, p = pearsonr(df_g[var], df_g['자살률'])
        sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "n.s."
        gender_corr[gender][var] = {'r': r, 'p': p, 'sig': sig}
        print(f"{var:<15} | {r:>12.4f} | {p:>12.4e} | {sig:<8}")

# C. Correlation by Year (Overall gender '계')
year_corr = {}
print("\n[Correlation by Year (Overall gender '계')]")
df_reg_total = df_regional[df_regional['성별'] == '계']
for year in sorted(df_reg_total['연도'].unique()):
    print(f"\n연도: {year}")
    print(f"{'변수':<15} | {'상관계수 (r)':<12} | {'p-value':<12} | {'유의성':<8}")
    print("-" * 55)
    year_corr[year] = {}
    df_y = df_reg_total[df_reg_total['연도'] == year]
    for var in satisfaction_variables:
        r, p = pearsonr(df_y[var], df_y['자살률'])
        sig = "***" if p < 0.001 else "**" if p < 0.01 else "*" if p < 0.05 else "n.s."
        year_corr[year][var] = {'r': r, 'p': p, 'sig': sig}
        print(f"{var:<15} | {r:>12.4f} | {p:>12.4e} | {sig:<8}")

# Save correlation text results for reference
with open(os.path.join(artifact_dir, "correlation_results.txt"), "w", encoding="utf-8") as f:
    f.write("=== OVERALL CORRELATION ===\n")
    for k, v in overall_corr.items():
        f.write(f"{k}: r={v['r']:.4f}, p={v['p']:.4e} ({v['sig']})\n")
    f.write("\n=== CORRELATION BY GENDER ===\n")
    for g, metrics in gender_corr.items():
        f.write(f"Gender: {g}\n")
        for k, v in metrics.items():
            f.write(f"  {k}: r={v['r']:.4f}, p={v['p']:.4e} ({v['sig']})\n")
    f.write("\n=== CORRELATION BY YEAR (GENDER: 계) ===\n")
    for y, metrics in year_corr.items():
        f.write(f"Year: {y}\n")
        for k, v in metrics.items():
            f.write(f"  {k}: r={v['r']:.4f}, p={v['p']:.4e} ({v['sig']})\n")

# ==========================================
# 3. VISUALIZATION GENERATION
# ==========================================
print("\nGenerating charts...")

# Chart 1: Heatmap of all satisfaction variables vs Suicide Rate (on regional data)
plt.figure(figsize=(10, 8))
# Compute correlation matrix between these variables
corr_cols = satisfaction_variables + ['자살률']
corr_matrix = df_regional[corr_cols].corr()
sns.heatmap(corr_matrix, annot=True, cmap='coolwarm', fmt=".3f", linewidths=0.5, annot_kws={"size": 11})
plt.title("삶의 만족도 하위 항목 및 자살률 간의 상관관계 히트맵 (시도별 데이터)", fontsize=14, pad=15)
plt.tight_layout()
heatmap_path = os.path.join(artifact_dir, "heatmap.png")
plt.savefig(heatmap_path, dpi=300)
plt.close()
print(f"Saved Heatmap to {heatmap_path}")

# Chart 2: Scatter plot with regression line (Overall Satisfaction vs Suicide Rate, gender '계')
df_reg_all_gender = df_regional[df_regional['성별'] == '계']
plt.figure(figsize=(8, 6))
sns.regplot(data=df_reg_all_gender, x='만족율', y='자살률', 
            scatter_kws={'alpha':0.7, 'color':'#2c3e50', 's':40}, 
            line_kws={'color':'#e74c3c', 'linewidth':2})
r_val = overall_corr['만족율']['r']
p_val = overall_corr['만족율']['p']
# Calculate correlation specifically for '계' gender
r_val_total, p_val_total = pearsonr(df_reg_all_gender['만족율'], df_reg_all_gender['자살률'])
plt.title("삶의 만족율과 자살률 산점도 및 추세선 (전체 성별, 17개 시도)", fontsize=13, pad=10)
plt.xlabel("삶의 만족율 (%) (매우 만족 + 약간 만족)", fontsize=11)
plt.ylabel("자살률 (인구 10만 명당 명)", fontsize=11)
plt.text(0.05, 0.05, f"Pearson r = {r_val_total:.3f}\np-value = {p_val_total:.4e}", 
         transform=plt.gca().transAxes, bbox=dict(facecolor='white', alpha=0.8, boxstyle='round,pad=0.5'), fontsize=10)
plt.tight_layout()
scatter_overall_path = os.path.join(artifact_dir, "scatter_overall.png")
plt.savefig(scatter_overall_path, dpi=300)
plt.close()
print(f"Saved Overall Scatter Plot to {scatter_overall_path}")

# Chart 3: Scatter plot with regression line by Gender (Male vs Female, excluding '계')
df_reg_genders = df_regional[df_regional['성별'] != '계']
plt.figure(figsize=(9, 7))
g = sns.lmplot(data=df_reg_genders, x='만족율', y='자살률', hue='성별', 
               palette={'남자': '#2980b9', '여자': '#e84393'},
               markers=['o', 's'], height=6, aspect=1.2, scatter_kws={'alpha': 0.6, 's':40},
               legend=False)
plt.title("성별 삶의 만족율과 자살률 산점도 및 추세선 (17개 시도)", fontsize=13, pad=15)
plt.xlabel("삶의 만족율 (%)", fontsize=11)
plt.ylabel("자살률 (인구 10만 명당 명)", fontsize=11)
# Add manual legend with correlation coefficients
r_male = gender_corr['남자']['만족율']['r']
r_female = gender_corr['여자']['만족율']['r']
plt.legend(title='성별 (추세선)', labels=[
    f'남성 (r = {r_male:.3f})',
    f'여성 (r = {r_female:.3f})'
], loc='upper right')
plt.tight_layout()
scatter_gender_path = os.path.join(artifact_dir, "scatter_by_gender.png")
plt.savefig(scatter_gender_path, dpi=300)
plt.close()
print(f"Saved Gender Scatter Plot to {scatter_gender_path}")

# Chart 4: Dual-axis yearly trends (National '전국' data, gender '계')
df_nat_total = df_national[df_national['성별'] == '계'].sort_values('연도')
fig, ax1 = plt.subplots(figsize=(8, 5))

color = '#1abc9c'
ax1.set_xlabel('연도', fontsize=11, labelpad=10)
ax1.set_ylabel('전국 평균 삶의 만족율 (%)', color=color, fontsize=11)
line1 = ax1.plot(df_nat_total['연도'], df_nat_total['만족율'], color=color, marker='o', linewidth=2.5, label='삶의 만족율 (%)')
ax1.tick_params(axis='y', labelcolor=color)
ax1.set_xticks(df_nat_total['연도'])

ax2 = ax1.twinx()  
color = '#e74c3c'
ax2.set_ylabel('전국 자살률 (인구 10만 명당 명)', color=color, fontsize=11)
line2 = ax2.plot(df_nat_total['연도'], df_nat_total['자살률'], color=color, marker='s', linestyle='--', linewidth=2.5, label='자살률 (명)')
ax2.tick_params(axis='y', labelcolor=color)

# Combine legends
lines = line1 + line2
labels = [l.get_label() for l in lines]
ax1.legend(lines, labels, loc='upper left')

plt.title("연도별 전국 평균 삶의 만족율과 자살률 변화 추이 (2020-2024)", fontsize=13, pad=15)
fig.tight_layout()
yearly_trends_path = os.path.join(artifact_dir, "yearly_trends.png")
plt.savefig(yearly_trends_path, dpi=300)
plt.close()
print(f"Saved Yearly Trends to {yearly_trends_path}")

# Chart 5: Regional Quadrant Chart (Average Satisfaction vs Average Suicide Rate over 5 years for 17 regions, gender '계')
df_reg_mean = df_reg_all_gender.groupby('행정구역').agg({
    '만족율': 'mean',
    '자살률': 'mean'
}).reset_index()

plt.figure(figsize=(10, 8))
sns.scatterplot(data=df_reg_mean, x='만족율', y='자살률', s=100, color='#8e44ad', edgecolor='black', alpha=0.8)

# Calculate averages for quadrant lines
mean_sat = df_reg_mean['만족율'].mean()
mean_sui = df_reg_mean['자살률'].mean()

plt.axvline(x=mean_sat, color='#7f8c8d', linestyle='--', linewidth=1.5)
plt.axhline(y=mean_sui, color='#7f8c8d', linestyle='--', linewidth=1.5)

# Annotate regions
for idx, row in df_reg_mean.iterrows():
    # Subtle offset to avoid label overlapping the dot
    plt.text(row['만족율'] + 0.15, row['자살률'] + 0.1, row['행정구역'], fontsize=9, weight='bold')

plt.title("시도별 5개년 평균 삶의 만족율 및 자살률 분포 사분면 차트 (2020-2024)", fontsize=13, pad=15)
plt.xlabel("5개년 평균 삶의 만족율 (%)", fontsize=11)
plt.ylabel("5개년 평균 자살률 (인구 10만 명당 명)", fontsize=11)

# Add Quadrant Labels
xlim = plt.xlim()
ylim = plt.ylim()

# Quadrant I (Top-Right): High Sat, High Suicide
plt.text(xlim[1] - 0.5, ylim[1] - 0.8, "1사분면 (대비군)\n만족도 높음 / 자살률 높음", 
         color='#d35400', fontsize=10, weight='bold', ha='right', va='top', alpha=0.7)
# Quadrant II (Top-Left): Low Sat, High Suicide (Vulnerable)
plt.text(xlim[0] + 0.5, ylim[1] - 0.8, "2사분면 (취약 지역)\n만족도 낮음 / 자살률 높음", 
         color='#c0392b', fontsize=10, weight='bold', ha='left', va='top', alpha=0.7)
# Quadrant III (Bottom-Left): Low Sat, Low Suicide
plt.text(xlim[0] + 0.5, ylim[0] + 0.5, "3사분면 (소극형)\n만족도 낮음 / 자살률 낮음", 
         color='#7f8c8d', fontsize=10, weight='bold', ha='left', va='bottom', alpha=0.7)
# Quadrant IV (Bottom-Right): High Sat, Low Suicide (Ideal)
plt.text(xlim[1] - 0.5, ylim[0] + 0.5, "4사분면 (선도 지역)\n만족도 높음 / 자살률 낮음", 
         color='#27ae60', fontsize=10, weight='bold', ha='right', va='bottom', alpha=0.7)

plt.tight_layout()
regional_quadrant_path = os.path.join(artifact_dir, "regional_quadrant.png")
plt.savefig(regional_quadrant_path, dpi=300)
plt.close()
print(f"Saved Regional Quadrant Chart to {regional_quadrant_path}")

print("\nAnalysis and visualization generation complete!")
