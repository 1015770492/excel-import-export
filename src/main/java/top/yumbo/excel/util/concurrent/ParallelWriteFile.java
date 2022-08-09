package top.yumbo.excel.util.concurrent;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.RandomAccessFile;
import java.nio.MappedByteBuffer;
import java.nio.channels.FileChannel;
import java.util.concurrent.LinkedBlockingDeque;
import java.util.concurrent.ThreadPoolExecutor;
import java.util.concurrent.TimeUnit;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * @author jinhua
 * @date 2021/8/6 15:27
 * 并行写文件
 */
public class ParallelWriteFile {

    static ThreadPoolExecutor poll = new ThreadPoolExecutor(5, 5, 0, TimeUnit.MILLISECONDS, new LinkedBlockingDeque<>(10));

    public AtomicInteger index = new AtomicInteger(0);
    public long size;

    public int buffer_size;

    RandomAccessFile origin_raf;
    RandomAccessFile target_raf;
    FileChannel originFileChannel;
    FileChannel targetFileChannel;

    public ParallelWriteFile(File origin, int buffer_size) throws FileNotFoundException {
        RandomAccessFile origin_raf = new RandomAccessFile(origin, "r");
        this.originFileChannel = origin_raf.getChannel();
        this.buffer_size = buffer_size;
        this.size = origin.length();
    }


    public static void main(String[] args) throws FileNotFoundException {

        long start = System.currentTimeMillis();
        System.out.println("起始时间"+start+"ms");
        File orign = new File("src/test/resources/excel/ImportForQuarter_big/xl/worksheets/sheet1.xml");
        final ParallelWriteFile pf = new ParallelWriteFile(orign, 3000);

        for (int i = 0; i < 5; i++) {
            poll.execute(pf.new ReadTask());
        }
        long end = System.currentTimeMillis();
        System.out.println("耗时：" + (end - start) + "ms");


    }
    class ReadTask implements Runnable {
        @Override
        public void run() {
            // 这个任务中需要使用cas获取到当前的 index，并且读取index+buffer值，然后将index改为
            int cur_index;
            System.out.println("执行");
            while ((cur_index = index.get()) < size) {

                int target_index = (cur_index + buffer_size) > size ? (int) size : cur_index + buffer_size;

                if (index.compareAndSet(cur_index, target_index + 1)) {
                    //如果cas 成功就进行读写操作

//                    System.out.println(Thread.currentThread().getName() + "读取了cur_index" + cur_index + "到" + target_index + "的缓冲数据");
                    byte[] content = readFile(cur_index, target_index);
                    String str = new String(content);
                    System.out.println(str);
                    // 读取到了内容可以在下面进行一些别的处理操作

                }
            }
            System.out.println("========="+Thread.currentThread()+"执行结束,"+System.currentTimeMillis()+"ms");

        }

        public byte[] readFile(int start_index, int end_index) {

            // 读取文件,使用一个map内存映射进行读取，可以加速读取吧
            MappedByteBuffer map;
            byte[] byteArr = new byte[end_index - start_index];
            try {
                map = originFileChannel.map(FileChannel.MapMode.READ_ONLY, start_index, end_index - start_index);

                map.get(byteArr, 0, end_index - start_index);
            } catch (Exception e) {
                System.out.println("读取" + start_index + "到" + end_index + "失败");
                e.printStackTrace();
            }
            return byteArr;
        }

    }

    class WriteTask implements Runnable {

        @Override
        public void run() {

            byte[] a = new byte[1024];
            int cur_index;
//            System.out.println("执行");
            while ((cur_index = index.get()) < size) {

                int target_index = (cur_index + buffer_size) > size ? (int) size : cur_index + buffer_size;
                if (index.compareAndSet(cur_index, target_index + 1)) {
                    //如果cas 成功就进行读写操作
                    //成功
                    System.out.println(Thread.currentThread().getName() + "写了cur_index" + cur_index + "到" + target_index + "的缓冲数据");
                    writeFile(cur_index, target_index, a);
                    // 读取
                }
            }
        }

        public void writeFile(int orign_index, int target_index, byte[] content) {
            //然后进行写
            // 读取文件,使用一个map内存映射进行读取，可以加速读取吧
            MappedByteBuffer map;
            byte[] byteArr = new byte[target_index - orign_index];
            try {
                map = targetFileChannel.map(FileChannel.MapMode.READ_ONLY, orign_index, target_index - orign_index);

                map.position();
                map.put(content);
            } catch (IOException e) {
                e.printStackTrace();
            }


        }
    }




}

