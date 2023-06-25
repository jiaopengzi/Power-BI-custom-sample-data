# 179\_自动生成 千万级 Power BI 示例数据

在早一些是时候，我曾写过一个示例数据[《赠送300家门店260亿销售额的零售企业Power BI实战示例数据》](https://jiaopengzi.com/1435.html)，本次我们对该示例数据做了一些调整。



## 一、更新内容

1. 针对有一些朋友不会使用 vba 模块，我们增加了 UI 操作。

    ![图-01](https://image.jiaopengzi.com/blog/202306251058119.png)

    

    填写几个简单的参数即可生成相应 Power BI 示例数据

    ![图-02](https://image.jiaopengzi.com/blog/202306251046880.png)

    

2. 丰富了表格内容。

    ![图-03](https://image.jiaopengzi.com/blog/202306251050320.png)

    在原来表格的基础上增加了：`D10_组织表`、`T11_门店表_租赁`、`T12_门店表_装修`、`T60_员工信息表`、`T61_人工成本表`。

    **需要注意本次的更新和原来的字段和表格名称都有所变化，不兼容。**



## 二、使用 vba 模块。

### 1、步骤

1. 新建 access 文件

2. 打开 access 

3. 打开 visual basic 窗口

4. 右键导入源码中的 `Power-BI-custom-sample-data.bas`

5. 在 **main** 函数中配置好 vba 中的参数

```vb
productQuantity = 200       '产品数量；建议ShopQuantity∈[7,1688]。
ShopQuantity = 5            '门店数量；建议ShopQuantity∈[1,390]。
MaxInventoryDays = 14       '入库间隔最大数；建议ShopQuantity∈[5,20]。
```

6. 运行，等待数据模拟完成即可。



### 2、关于运行时间

#### 电脑配置

- CPU：12th Gen Intel(R) Core(TM) i9-12900KF   3.20 GHz
- 内存：RAM 32.0 GB

#### 运行时间参考

- 如上电脑配置 + ShopQuantity=5   的配置：大约需要   20 秒，每秒按照业务逻辑生成约 1万行+的数据；生成   20 万行+ demo数据。

- 如上电脑配置 + ShopQuantity=10  的配置：大约需要   60 秒，每秒按照业务逻辑生成约 1万行+的数据；生成   60 万行+ demo数据。

- 如上电脑配置 + ShopQuantity=100 的配置：大约需要  350 秒，每秒按照业务逻辑生成约 1万行+的数据；生成  360 万行+ demo数据。

- 如上电脑配置 + ShopQuantity=300 的配置：大约需要 1000 秒，每秒按照业务逻辑生成约 1万行+的数据；生成 1000 万行+ demo数据。



![图-04](https://image.jiaopengzi.com/blog/202306251147726.png)



基本满足实战学习所用，可以根据自己需要调节数量,`0`结尾的包含运行的窗体。



## 三、特殊表格数据清洗

原始数据按照合同信息记录，只有开始日期，结束日期，还有按照年度增长的比例。

![图-05](https://image.jiaopengzi.com/blog/202306251127611.png)



经过 Power Query 清洗后，数据拆分到天，便于后续的`DAX`建模

![图-06](https://image.jiaopengzi.com/blog/202306251129556.png)



诸如此类表格，还有:`T12_门店表_装修`、`T50_销售目标表`、`T64_人工成本表`,我们将在直播中讲解。



## 四、数据加载到 Power BI

表间关系

![图-07](https://image.jiaopengzi.com/blog/202306251121711.png)



**后续我们将在此示例数据基础上展开更过的学习和探索，欢迎加入焦棚子的会员。**



## 直播预告

B站|微信视频号 同步直播。

**时间：2023年6月26日 晚 20:00**

![图-08](https://image.jiaopengzi.com/wp-content/uploads/bilibili-live-qr-code.png)



### 附件下载

[https://jiaopengzi.com/3011.html](https://jiaopengzi.com/3011.html)



------

### 请关注

全网同名搜索 **焦棚子**

如果对你有帮助，请 **点赞**、**关注**、**三连** 支持一下，这是我们更新的动力。

![图-09](https://image.jiaopengzi.com/wp-content/uploads/jiaopengzierweima.png)



by 焦棚子
