package top.yumbo.excel.util.concurrent;

import java.lang.management.ManagementFactory;
import java.lang.management.MemoryMXBean;
import java.lang.management.ThreadInfo;
import java.lang.management.ThreadMXBean;

/**
 * @author jinhua
 * @date 2022/7/27 22:30
 */
public class MemoryController {

    public static void main(String[] args) {
        MemoryMXBean memoryMXBean = ManagementFactory.getMemoryMXBean();
        ThreadMXBean threadMXBean = ManagementFactory.getThreadMXBean();
        for (long id : threadMXBean.getAllThreadIds()) {
            ThreadInfo threadInfo = threadMXBean.getThreadInfo(id);
//            threadInfo.get
        }
        long maxMemory = memoryMXBean.getHeapMemoryUsage().getMax();
        // 获取已提交的内存
        long committedMemory = memoryMXBean.getHeapMemoryUsage().getCommitted();
        // 获取已使用的内存
        long usedMemory = memoryMXBean.getHeapMemoryUsage().getUsed();
        double usedMemoryPercentage = (double) usedMemory / maxMemory * 100;
        System.out.println(usedMemoryPercentage);
    }
}
