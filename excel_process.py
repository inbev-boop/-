import pandas as pd
import pickle
from sklearn.feature_extraction.text import TfidfVectorizer
from sklearn.preprocessing import StandardScaler
import numpy as np


def process_excel(input_excel_path, output_excel_path, model_path=r'D:\PyCharm Community Edition 2024.1.6\TEST\ML\K_model.pkl'):
    # 1. 加载模型
    with open(model_path, 'rb') as f:
        data = pickle.load(f)
        model = data["model"]
        tfidf = data["tfidf"]
        scaler = data["scaler"]



    # 2. 读取Excel文件
    df = pd.read_excel(input_excel_path)

    # 3. 检查必要的列是否存在
    required_columns = ['name', '是否喜欢安静', '是否早睡',"是否抽烟", "是否很注重卫生", "爱好"]
    if not set(required_columns).issubset(df.columns):
        missing = set(required_columns) - set(df.columns)
        raise ValueError(f"Excel文件缺少必要的列: {missing}")


    structured_features = df[['是否早睡', '是否喜欢安静', '是否抽烟', '是否很注重卫生']].astype(int)
    text_features = df['爱好'].astype(str)

    # 5. 处理文本特征
    text_vectors = tfidf.transform(text_features)

    # 6. 标准化结构化特征
    structured_scaled = scaler.fit_transform(structured_features)

    # 7. 融合特征
    X = np.hstack([structured_scaled, text_vectors.toarray()])

    # 8. 进行预测
    df['cluster'] = model.predict(X)

    # 5. 保存结果到新的Excel文件
    df.to_excel(output_excel_path, index=False)



# 示例用法
if __name__ == "__main__":
    process_excel(
        input_excel_path=r"D:\PyCharm Community Edition 2024.1.6\TEST\ML\tree.xlsx",
        output_excel_path=r'D:\PyCharm Community Edition 2024.1.6\TEST\ML\output.xlsx'
    )
