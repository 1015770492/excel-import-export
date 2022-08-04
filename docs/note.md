
### 2022/07/31
从版本1.3.21开始开始记录下一个开发版本需要解决的问题

1. poi本身的内存问题导致OOM (重要)
2. 根据Consumer<List<ImportForQuarter>> consumer = (list) -> {} 获取ImportForQuarter.class，可以减少传一个参数
3. 需要重新整理1.3.21 的文档
4. 引入日志，使用日志代替控制台输出日志
5. 达到增量导入的效果，重跑程序，不需要重新执行 ExcelImportExportUtils.importExcelConsumer 方法，解决海量数据问题
6. 优化代码，提高效率
7. 提供与DB（可以是关系型数据库，也可以是非关系型数据库）交互的字典映射。

对于问题5，我使用的是forkjoin框架的思想，将任务拆分成ln个job，有m个线程并行执行这n个任务。

也就是说处理excel进行consumer消费的时候可能是随机的消费了 x行数据。

出现error后我希望是，当我重新执行的时候，不希望重新消费这x行数据，而只需要消费剩下的数据。避免重复执行

