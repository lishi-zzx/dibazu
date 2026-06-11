import pandas as pd
import matplotlib.pyplot as plt
from sklearn.model_selection import train_test_split
from sklearn.linear_model import LogisticRegression
from sklearn.preprocessing import StandardScaler
from sklearn.metrics import (
    accuracy_score, classification_report, roc_auc_score,
    roc_curve, confusion_matrix, ConfusionMatrixDisplay
)
from imblearn.over_sampling import SMOTE  # 处理不平衡数据

# ===================== 1. 读取数据 =====================
df = pd.read_excel(r'C:\Users\TRYX\Desktop\diabetic_data - 副本.xlsx')

# ===================== 2. 划分特征X 和 标签y =====================
X = df.drop('has_complication', axis=1)
y = df['has_complication']

# ===================== 3. 划分训练集 / 测试集 =====================
X_train, X_test, y_train, y_test = train_test_split(
    X, y, test_size=0.3, random_state=42, stratify=y
)

# ===================== 4. 特征标准化 =====================
scaler = StandardScaler()
X_train_scaled = scaler.fit_transform(X_train)
X_test_scaled = scaler.transform(X_test)

# ===================== 5. 处理数据不平衡（关键！） =====================
smote = SMOTE(random_state=42)
X_train_balanced, y_train_balanced = smote.fit_resample(X_train_scaled, y_train)

# ===================== 6. 训练 Logistic 回归模型 =====================
model = LogisticRegression(random_state=42, max_iter=1000)
model.fit(X_train_balanced, y_train_balanced)
print("✅ Logistic 回归模型训练完成！")

# ===================== 7. 模型预测 & 评估 =====================
y_pred = model.predict(X_test_scaled)
y_pred_proba = model.predict_proba(X_test_scaled)[:, 1]

print("\n===== 模型性能评估 =====")
print(f"准确率：{accuracy_score(y_test, y_pred):.4f}")
print(f"AUC值：{roc_auc_score(y_test, y_pred_proba):.4f}")
print("\n分类报告：")
print(classification_report(y_test, y_pred))

# ===================== 8. 特征影响分析（可直接写报告） =====================
coef_df = pd.DataFrame({
    '特征名': X.columns,
    '回归系数': model.coef_[0],
    '影响方向': ['风险增加' if c > 0 else '风险降低' for c in model.coef_[0]],
    '影响强度': abs(model.coef_[0])
}).sort_values(by='影响强度', ascending=False)

print("\n===== 特征对并发症的影响排名 =====")
print(coef_df.to_string(index=False))

# ===================== 9. 可视化图表（自动保存到桌面） =====================
plt.rcParams['font.sans-serif'] = ['SimHei']
plt.rcParams['axes.unicode_minus'] = False

# 1. 特征重要性图
plt.figure(figsize=(10, 6))
top_features = coef_df.head(10)
plt.barh(top_features['特征名'][::-1], top_features['影响强度'][::-1], color='#4472C4')
plt.xlabel('影响强度')
plt.title('糖尿病并发症 影响因素TOP10')
plt.tight_layout()
plt.savefig(r'C:\Users\TRYX\Desktop\特征重要性.png', dpi=300)
plt.show()

# 2. ROC曲线
fpr, tpr, _ = roc_curve(y_test, y_pred_proba)
plt.figure(figsize=(6,6))
plt.plot(fpr, tpr, label=f'AUC = {roc_auc_score(y_test,y_pred_proba):.3f}')
plt.plot([0,1],[0,1],'--')
plt.xlabel('假阳性率')
plt.ylabel('真阳性率')
plt.title('ROC曲线')
plt.legend()
plt.tight_layout()
plt.savefig(r'C:\Users\TRYX\Desktop\ROC曲线.png', dpi=300)
plt.show()

# 3. 混淆矩阵
cm = confusion_matrix(y_test, y_pred)
ConfusionMatrixDisplay(cm, display_labels=['无并发症','有并发症']).plot(cmap='Blues')
plt.title('混淆矩阵')
plt.tight_layout()
plt.savefig(r'C:\Users\TRYX\Desktop\混淆矩阵.png', dpi=300)
plt.show()

print("\n🎉 全部完成！图表已保存到桌面！")
import numpy as np
import matplotlib.pyplot as plt

# ---------------------- 1. 选择两个最重要的特征 ----------------------
# 这里用影响强度前两名的特征：诊断总数、住院次数
feature1_name = '诊断总数'
feature2_name = '住院次数'

# 获取这两个特征在标准化数据中的索引
idx1 = X.columns.get_loc(feature1_name)
idx2 = X.columns.get_loc(feature2_name)

# 提取这两个特征的训练集和测试集数据
X_train_2d = X_train_scaled[:, [idx1, idx2]]
X_test_2d = X_test_scaled[:, [idx1, idx2]]

# 用这两个特征重新训练一个Logistic回归模型（为了画决策边界）
model_2d = LogisticRegression(random_state=42, max_iter=1000)
model_2d.fit(X_train_2d, y_train)

# ---------------------- 2. 定义绘制决策边界的函数 ----------------------
def plot_decision_boundary(model, X, y, title, save_path):
    plt.rcParams['font.sans-serif'] = ['SimHei']
    plt.rcParams['axes.unicode_minus'] = False
    
    # 设置网格范围
    h = 0.02
    x_min, x_max = X[:, 0].min() - 1, X[:, 0].max() + 1
    y_min, y_max = X[:, 1].min() - 1, X[:, 1].max() + 1
    xx, yy = np.meshgrid(np.arange(x_min, x_max, h),
                         np.arange(y_min, y_max, h))
    
    # 预测网格上的点
    Z = model.predict(np.c_[xx.ravel(), yy.ravel()])
    Z = Z.reshape(xx.shape)
    
    # 绘制决策边界背景
    plt.figure(figsize=(6, 5))
    plt.contourf(xx, yy, Z, cmap=plt.cm.RdYlGn, alpha=0.6)
    
    # 绘制样本点
    scatter = plt.scatter(X[:, 0], X[:, 1], c=y, cmap=plt.cm.RdYlGn, edgecolors='k', s=30)
    plt.xlabel(feature1_name + '（标准化）')
    plt.ylabel(feature2_name + '（标准化）')
    plt.title(title)
    plt.legend(handles=scatter.legend_elements()[0], labels=['无并发症', '有并发症'])
    plt.tight_layout()
    plt.savefig(save_path, dpi=300)
    plt.show()

# ---------------------- 3. 分别绘制训练集和测试集的图 ----------------------
# 训练集
plot_decision_boundary(
    model_2d, X_train_2d, y_train,
    title='LOGISTIC（Training set）',
    save_path=r'C:\Users\TRYX\Desktop\训练集决策边界.png'
)

# 测试集
plot_decision_boundary(
    model_2d, X_test_2d, y_test,
    title='LOGISTIC（Test set）',
    save_path=r'C:\Users\TRYX\Desktop\测试集决策边界.png'
)
