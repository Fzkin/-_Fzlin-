## 第一题
空值处理，部分数据求和，对下一个月份客户类型数量求平均值，使用多元线性回归模型进行预测下一个月的总税收。

## 第二题
循环为日期，按照日期顺序进行顺次循环
因为数据中客户使用商品的日期不同，需求量不同，进行数据标记，随着日期循环，逐个递减，当Flag为0时，即正常发放物资。


## 总结：
第一题，需要对客户类型模型进行预测，不能简单地求平均值预测。
第二题，提出的假设有点不切实际，3天内到达的物资，全部假设为最后一天到达，有点太理想化，要求在保证正常运转的情况下，求出最极限的情况，理想情况太简单了。
