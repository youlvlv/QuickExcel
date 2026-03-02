package com.lizhiwei.quickExcel.test;

import com.lizhiwei.quickExcel.config.ExcelConfig;
import com.lizhiwei.quickExcel.util.ReadExcel;

import java.io.File;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;
import java.util.UUID;

import java.io.*;
import java.lang.management.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;

public class Main {


    /**
     * 内存监控器 - 单文件版本
     * 用法：java MemoryMonitor [测试时长秒数] [采样间隔毫秒]
     * 示例：java MemoryMonitor 60 1000
     */
    public static class MemoryMonitor {

        private static final DateTimeFormatter TIME_FORMAT =
                DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

        private final String csvFile;
        private final long intervalMs;
        private final ScheduledExecutorService scheduler =
                Executors.newSingleThreadScheduledExecutor();
        private final List<String> records = new CopyOnWriteArrayList<>();
        private volatile boolean running = false;

        public MemoryMonitor(String csvFile, long intervalMs) {
            this.csvFile = csvFile;
            this.intervalMs = intervalMs;
        }

        /** 开始监控 */
        public void start() {
            if (running) return;
            running = true;

            // 写入 CSV 表头
            writeHeader();

            // 定时采集
            scheduler.scheduleAtFixedRate(() -> {
                try {
                    String record = capture();
                    records.add(record);
                    appendToFile(record);
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }, 0, intervalMs, TimeUnit.MILLISECONDS);

            System.out.println("🔍 内存监控已启动，间隔：" + intervalMs + "ms");
        }

        /** 停止监控 */
        public void stop() {
            if (!running) return;
            running = false;
            scheduler.shutdown();
            try {
                scheduler.awaitTermination(5, TimeUnit.SECONDS);
            } catch (InterruptedException e) {
                scheduler.shutdownNow();
            }
            System.out.println("✅ 监控停止，共采集 " + records.size() + " 条数据");
            System.out.println("📁 数据文件：" + csvFile);
        }

        /** 采集单次内存快照 */
        private String capture() {
            Runtime runtime = Runtime.getRuntime();
            MemoryMXBean memoryBean = ManagementFactory.getMemoryMXBean();
            MemoryUsage heap = memoryBean.getHeapMemoryUsage();
            MemoryUsage nonHeap = memoryBean.getNonHeapMemoryUsage();

            long timestamp = System.currentTimeMillis();
            String timeStr = LocalDateTime.now().format(TIME_FORMAT);

            // 计算各项内存 (MB)
            double heapUsed = toMB(heap.getUsed());
            double heapCommitted = toMB(heap.getCommitted());
            double heapMax = toMB(heap.getMax());
            double nonHeapUsed = toMB(nonHeap.getUsed());
            double totalUsed = toMB(runtime.totalMemory() - runtime.freeMemory());
            double maxMemory = toMB(runtime.maxMemory());

            // GC 统计
            long gcCount = 0, gcTime = 0;
            for (GarbageCollectorMXBean gc : ManagementFactory.getGarbageCollectorMXBeans()) {
                if (gc.getCollectionCount() >= 0) gcCount += gc.getCollectionCount();
                if (gc.getCollectionTime() >= 0) gcTime += gc.getCollectionTime();
            }

            // 线程数
            int threadCount = Thread.activeCount();

            return String.format("%s,%d,%.2f,%.2f,%.2f,%.2f,%.2f,%.2f,%d,%d,%d",
                    timeStr, timestamp, heapUsed, heapCommitted, heapMax,
                    nonHeapUsed, totalUsed, maxMemory, gcCount, gcTime, threadCount);
        }

        private double toMB(long bytes) {
            return bytes / 1024.0 / 1024.0;
        }

        private void writeHeader() {
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFile))) {
                writer.write("timestamp,timestampMs,heapUsedMB,heapCommittedMB,heapMaxMB," +
                        "nonHeapUsedMB,totalUsedMB,maxMemoryMB,gcCount,gcTimeMs,threadCount\n");
            } catch (IOException e) {
                throw new RuntimeException("写入表头失败", e);
            }
        }

        private void appendToFile(String record) {
            try (BufferedWriter writer = new BufferedWriter(new FileWriter(csvFile, true))) {
                writer.write(record + "\n");
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        /** 生成简单文本图表（无需 Python） */
        public void printTextChart() {
            if (records.isEmpty()) return;

            System.out.println("\n" + "=".repeat(60));
            System.out.println("📊 内存使用趋势 (文本图表)");
            System.out.println("=".repeat(60));

            // 找出最大值用于缩放
            double maxHeap = 0;
            for (String r : records) {
                String[] parts = r.split(",");
                double heapUsed = Double.parseDouble(parts[2]);
                if (heapUsed > maxHeap) maxHeap = heapUsed;
            }

            // 采样显示（最多 20 行）
            int step = Math.max(1, records.size() / 20);
            int barWidth = 50;

            for (int i = 0; i < records.size(); i += step) {
                String[] parts = records.get(i).split(",");
                String time = parts[0].substring(11, 19); // 只取时分秒
                double heapUsed = Double.parseDouble(parts[2]);
                int barLength = (int) (heapUsed / maxHeap * barWidth);

                System.out.printf("%s |%s %.1fMB\n",
                        time,
                        "█".repeat(barLength),
                        heapUsed);
            }
            System.out.println("=".repeat(60));
        }

        // ==================== 主方法 ====================
        public static void main(String[] args) throws InterruptedException {
            // 参数解析
            int durationSec = args.length > 0 ? Integer.parseInt(args[0]) : 10;
            long intervalMs = args.length > 1 ? Long.parseLong(args[1]) : 1;
            String csvFile = "memory_data.csv";

            System.out.println("🚀 内存监控启动");
            System.out.println("   测试时长：" + durationSec + "秒");
            System.out.println("   采样间隔：" + intervalMs + "ms");
            System.out.println("   输出文件：" + csvFile);
            System.out.println("-".repeat(60));

            // 启动监控
            MemoryMonitor monitor = new MemoryMonitor(csvFile, intervalMs);
            monitor.start();

            // 🔥🔥🔥 在这里执行你的工具类测试代码 🔥🔥🔥
            testYourTool();
            Thread.sleep(durationSec * 1000L);

            // 停止监控
            monitor.stop();

            // 打印文本图表
            monitor.printTextChart();

            // 打印 Python 绘图提示
            System.out.println("\n📌 生成精美图表，请执行:");
            System.out.println("   python plot_memory.py " + csvFile);
        }

        // ==================== 你的测试代码放这里 ====================
        private static void testYourTool() {
            System.out.println("🧪 开始执行工具类测试...");

            // 示例：模拟内存波动（替换为你的真实代码）
            new Thread(() -> {
                ExcelConfig.setDefaultReadEngine(ExcelConfig.ReadEngine.V2);
                ExcelConfig.createExcelImageSaveFunction((inputStream, fileExtension) -> {
                    // 保存文件
                    UUID uuid = UUID.randomUUID();
                    File file = new File("/Users/lizhiwei/Downloads/cache/" + uuid + "." + fileExtension);
                    Path path = Paths.get(file.toURI());
                    try {
                        Files.copy(inputStream, path, StandardCopyOption.REPLACE_EXISTING);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    return file.getPath();
                });
                var cache = ReadExcel.readExcelByV3(new File("/Users/lizhiwei/Downloads/典型案例.xlsx"),1,0,0,
                        DefectEntity.class);
                System.out.println(cache.size());
            }).start();
        }
    }
    public static void main(String[] args) throws InterruptedException {
        MemoryMonitor.main(args);
    }
}
