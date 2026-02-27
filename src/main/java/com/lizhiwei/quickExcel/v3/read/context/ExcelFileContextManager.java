package com.lizhiwei.quickExcel.v3.read.context;

import java.io.File;
import java.lang.ref.WeakReference;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

/**
 * Excel 文件上下文管理器
 * 单例模式，管理文件上下文的缓存和生命周期
 */
public class ExcelFileContextManager {
    
    private static volatile ExcelFileContextManager instance;
    
    // 文件上下文缓存（key -> 上下文）
    // 使用 ConcurrentHashMap 保证线程安全
    private final Map<String, ExcelFileContext> contextCache;
    
    // 使用 WeakReference 包装，允许 GC 在内存紧张时回收
    private final Map<String, WeakReference<ExcelFileContext>> weakContextCache;

    private final ThreadLocal<ExcelFileContext> threadContext = new ThreadLocal<>();
    
    // 最大缓存数量
    private static final int MAX_CACHE_SIZE = 100;
    
    // 是否使用弱引用（默认 true，内存紧张时自动释放）
    private boolean useWeakReference = true;
    
    // 是否自动清理过期缓存（文件被修改）
    private boolean autoCleanModified = true;
    
    private ExcelFileContextManager() {
        this.contextCache = new ConcurrentHashMap<>();
        this.weakContextCache = new ConcurrentHashMap<>();
    }
    
    /**
     * 获取单例实例
     */
    public static ExcelFileContextManager getInstance() {
        if (instance == null) {
            synchronized (ExcelFileContextManager.class) {
                if (instance == null) {
                    instance = new ExcelFileContextManager();
                }
            }
        }
        return instance;
    }
    
    /**
     * 获取文件上下文（如果不存在则创建）
     * @param file Excel 文件
     * @return 文件上下文
     */
    public ExcelFileContext getContext(File file) {
        String key = generateKey(file);
        
        // 检查缓存
        ExcelFileContext context = getFromCache(key);
        
        if (context != null) {
            // 检查文件是否被修改
            if (autoCleanModified && context.isModified()) {
                // 文件已修改，移除旧缓存
                removeContext(key);
                context = null;
            }
        }
        
        if (context == null) {
            // 创建新上下文
            context = new ExcelFileContext(file);
            putToCache(key, context);
        }

        if (threadContext.get() == null) {
            threadContext.set(context);
        }
        
        return context;
    }
    
    /**
     * 获取文件上下文（按文件路径）
     * @param filePath 文件路径
     * @return 文件上下文，如果文件不存在返回 null
     */
    public ExcelFileContext getContext(String filePath) {
        File file = new File(filePath);
        if (!file.exists()) {
            return null;
        }
        return getContext(file);
    }
    
    /**
     * 移除文件上下文
     * @param file Excel 文件
     */
    public void removeContext(File file) {
        threadContext.remove();
        removeContext(generateKey(file));
    }
    
    /**
     * 移除文件上下文
     * @param key 缓存键
     */
    public void removeContext(String key) {
        ExcelFileContext context = contextCache.remove(key);
        weakContextCache.remove(key);
        
        if (context != null) {
            try {
                context.close();
            } catch (Exception e) {
                // 忽略关闭异常
            }
        }
    }
    
    /**
     * 清理所有缓存
     */
    public void clearAll() {
        // 关闭所有上下文
        for (ExcelFileContext context : contextCache.values()) {
            try {
                context.close();
            } catch (Exception e) {
                // 忽略
            }
        }
        
        contextCache.clear();
        weakContextCache.clear();
    }
    
    /**
     * 清理过期/无效的缓存
     */
    public void cleanUp() {
        // 清理被修改的文件
        if (autoCleanModified) {
            contextCache.entrySet().removeIf(entry -> {
                if (entry.getValue().isModified()) {
                    try {
                        entry.getValue().close();
                    } catch (Exception e) {
                        // 忽略
                    }
                    return true;
                }
                return false;
            });
        }
        
        // 清理弱引用中已被 GC 的对象
        weakContextCache.entrySet().removeIf(entry -> entry.getValue().get() == null);
        
        // 如果缓存过多，清理最久未使用的（简单策略：清理一半）
        if (contextCache.size() > MAX_CACHE_SIZE) {
            int toRemove = contextCache.size() / 2;
            int removed = 0;
            for (String key : contextCache.keySet()) {
                if (removed >= toRemove) break;
                removeContext(key);
                removed++;
            }
        }
    }
    
    /**
     * 获取缓存大小
     */
    public int getCacheSize() {
        return contextCache.size();
    }
    
    /**
     * 设置是否使用弱引用
     */
    public void setUseWeakReference(boolean useWeakReference) {
        this.useWeakReference = useWeakReference;
    }
    
    /**
     * 设置是否自动清理被修改的文件缓存
     */
    public void setAutoCleanModified(boolean autoCleanModified) {
        this.autoCleanModified = autoCleanModified;
    }

    public ThreadLocal<ExcelFileContext> getThreadContext() {
        return threadContext;
    }

    /**
     * 生成缓存键
     */
    private String generateKey(File file) {
        // 使用绝对路径 + 修改时间作为 key
        return file.getAbsolutePath() + "@" + file.lastModified();
    }
    
    /**
     * 从缓存获取
     */
    private ExcelFileContext getFromCache(String key) {
        // 先尝试强引用缓存
        ExcelFileContext context = contextCache.get(key);
        if (context != null) {
            return context;
        }
        
        // 尝试弱引用缓存
        if (useWeakReference) {
            WeakReference<ExcelFileContext> ref = weakContextCache.get(key);
            if (ref != null) {
                context = ref.get();
                if (context != null) {
                    // 重新放入强引用缓存
                    contextCache.put(key, context);
                    return context;
                } else {
                    // 已被 GC，清理
                    weakContextCache.remove(key);
                }
            }
        }
        
        return null;
    }
    
    /**
     * 放入缓存
     */
    private void putToCache(String key, ExcelFileContext context) {
        contextCache.put(key, context);
        
        if (useWeakReference) {
            weakContextCache.put(key, new WeakReference<>(context));
        }
        
        // 定期清理
        if (contextCache.size() % 10 == 0) {
            cleanUp();
        }
    }
}
