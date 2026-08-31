"""
Week 4 - Mock-up visualization generator.
Produces illustrative charts (PNG) clearly labelled as NOT based on actual data.
"""
import os
import numpy as np
import matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator

OUT = r"C:\Users\Lenovo\Student_Performance_Prediction\Week4_Report\assets"
os.makedirs(OUT, exist_ok=True)

rng = np.random.default_rng(42)
BLUE = '#1A3C6E'
LBLUE = '#2C5F8A'
ORANGE = '#C0563C'
GREEN = '#2E7D5B'
GREY = '#4A4A4A'

def stamp(ax, x=0.98, y=0.015, text="Illustrative Example - Not Based on Actual Dataset"):
    ax.text(x, y, text, transform=ax.transAxes, fontsize=8, ha='right', va='bottom',
            style='italic', color='#8A8A8A')

plt.rcParams.update({'font.family': 'DejaVu Sans', 'axes.edgecolor': '#BBBBBB',
                     'axes.linewidth': 0.8})

# ---------- Viz 1: Study Hours vs Final Score (scatter) ----------
n = 200
study = rng.normal(5, 1.5, n)
score = 55 + 5*study + rng.normal(0, 6, n)
fig, ax = plt.subplots(figsize=(7, 4.5))
ax.scatter(study, score, s=28, alpha=0.65, color=LBLUE, edgecolor='white', linewidth=0.4)
m, b = np.polyfit(study, score, 1)
xs = np.linspace(study.min(), study.max(), 50)
ax.plot(xs, m*xs+b, color=ORANGE, lw=2, label='Potential trend line')
ax.set_xlabel('Study Hours (per week)')
ax.set_ylabel('Final Score')
ax.set_title('Study Hours vs Final Score\n(Illustrative relationship)', fontsize=11, color=BLUE)
ax.legend(frameon=False, fontsize=9)
ax.grid(alpha=0.25, lw=0.5)
stamp(ax)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'viz1_study_score.png'), dpi=160)
plt.close(fig)

# ---------- Viz 2: Attendance vs Final Score (scatter + trend) ----------
att = rng.uniform(60, 100, n)
score2 = 30 + 0.55*att + rng.normal(0, 5, n)
fig, ax = plt.subplots(figsize=(7, 4.5))
ax.scatter(att, score2, s=28, alpha=0.65, color=GREEN, edgecolor='white', linewidth=0.4)
m2, b2 = np.polyfit(att, score2, 1)
xs2 = np.linspace(att.min(), att.max(), 50)
ax.plot(xs2, m2*xs2+b2, color=ORANGE, lw=2, label='Potential trend line')
ax.set_xlabel('Attendance (%)')
ax.set_ylabel('Final Score')
ax.set_title('Attendance vs Final Score\n(Illustrative relationship)', fontsize=11, color=BLUE)
ax.legend(frameon=False, fontsize=9)
ax.grid(alpha=0.25, lw=0.5)
stamp(ax)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'viz2_att_score.png'), dpi=160)
plt.close(fig)

# ---------- Viz 3: Final Score Distribution (histogram) ----------
fig, ax = plt.subplots(figsize=(7, 4.5))
ax.hist(score2, bins=20, color=LBLUE, edgecolor='white', alpha=0.85)
ax.axvline(score2.mean(), color=ORANGE, lw=2, ls='--', label='Mean (illustrative)')
ax.set_xlabel('Final Score')
ax.set_ylabel('Number of Students')
ax.set_title('Final Score Distribution\n(Illustrative hypothetical distribution)', fontsize=11, color=BLUE)
ax.legend(frameon=False, fontsize=9)
ax.grid(alpha=0.25, lw=0.5, axis='y')
stamp(ax)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'viz3_score_dist.png'), dpi=160)
plt.close(fig)

# ---------- Viz 4: Performance by Student Group (box plot) ----------
groups = ['Low', 'Medium', 'High']
data = [score2[att<75], score2[(att>=75)&(att<90)], score2[att>=90]]
fig, ax = plt.subplots(figsize=(7, 4.5))
bp = ax.boxplot(data, tick_labels=groups, patch_artist=True, widths=0.55)
for patch, col in zip(bp['boxes'], [BLUE, LBLUE, GREEN]):
    patch.set_facecolor(col); patch.set_alpha(0.7)
for med in bp['medians']:
    med.set_color('black'); med.set_linewidth(1.5)
ax.set_xlabel('Attendance Group (illustrative categories)')
ax.set_ylabel('Final Score')
ax.set_title('Performance by Student Group\n(Illustrative comparison)', fontsize=11, color=BLUE)
ax.grid(alpha=0.25, lw=0.5, axis='y')
stamp(ax)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'viz4_group_box.png'), dpi=160)
plt.close(fig)

# ---------- Viz 5: Feature Correlation Heatmap ----------
labels = ['Study\nHours','Attendance','Prev\nScore','Absences','Age','Internet','Final\nScore']
rv = rng.normal(0,0.08,(7,7))
np.fill_diagonal(rv, 1)
# force plausible mock pattern (illustrative)
rv[0,-1]=0.55; rv[1,-1]=0.48; rv[2,-1]=0.60; rv[3,-1]=-0.35
rv[4,-1]=0.05
for i in range(7):
    for j in range(7):
        rv[i,j]=rv[j,i]
fig, ax = plt.subplots(figsize=(7, 5.2))
im = ax.imshow(rv, cmap='RdBu_r', vmin=-1, vmax=1)
ax.set_xticks(range(7)); ax.set_yticks(range(7))
ax.set_xticklabels(labels, fontsize=8); ax.set_yticklabels(labels, fontsize=8)
for i in range(7):
    for j in range(7):
        v = rv[i,j]
        ax.text(j, i, f'{v:.2f}', ha='center', va='center', fontsize=7,
                color='white' if abs(v)>0.45 else 'black')
cbar = fig.colorbar(im, ax=ax, fraction=0.046, pad=0.04)
cbar.set_label('Correlation', fontsize=9)
ax.set_title('Feature Correlation Heatmap\n(Illustrative - values are hypothetical)', fontsize=11, color=BLUE)
stamp(ax, y=0.02)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'viz5_heatmap.png'), dpi=160)
plt.close(fig)

# ---------- Viz 6: Actual vs Predicted (scatter) ----------
actual = rng.uniform(40, 100, n)
pred = actual + rng.normal(0, 6, n)
fig, ax = plt.subplots(figsize=(7, 4.5))
ax.scatter(actual, pred, s=28, alpha=0.6, color=LBLUE, edgecolor='white', linewidth=0.4)
lims = [35, 105]
ax.plot(lims, lims, color=ORANGE, lw=2, ls='--', label='Perfect prediction line')
ax.set_xlim(lims); ax.set_ylim(lims)
ax.set_xlabel('Actual Final Score')
ax.set_ylabel('Predicted Final Score')
ax.set_title('Actual vs Predicted Performance\n(Hypothetical - no model trained yet)', fontsize=10.5, color=BLUE)
ax.legend(frameon=False, fontsize=9, loc='lower right')
ax.grid(alpha=0.25, lw=0.5)
stamp(ax)
fig.tight_layout()
fig.savefig(os.path.join(OUT, 'viz6_actual_pred.png'), dpi=160)
plt.close(fig)

print('Charts generated:', sorted(os.listdir(OUT)))
