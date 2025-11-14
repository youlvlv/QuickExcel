package com.lizhiwei.quickExcel.util;

import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Method;
import java.lang.reflect.Proxy;

public class ConditionalProxyFactory {

    public static <T> T createProxyIfAvailable(
            Class<?> implClass,
            Class<T> expectedInterface) {


        try {
            // 4. 创建实例（要求无参构造）
            Object target = implClass.getDeclaredConstructor().newInstance();

            // 5. 创建代理
            Object proxy = Proxy.newProxyInstance(
                    expectedInterface.getClassLoader(),
                    new Class[]{expectedInterface},
                    new LoggingHandler(target)
            );

            return expectedInterface.cast(proxy);

        } catch (Exception e) {
            System.err.println("Failed to instantiate or proxy " + implClass.getName() + ": " + e.getMessage());
            return null;
        }
    }


    public static class LoggingHandler implements InvocationHandler {
        private final Object target;

        public LoggingHandler(Object target) {
            this.target = target;
        }

        @Override
        public Object invoke(Object proxy, Method method, Object[] args) throws Throwable {
            Object result = method.invoke(target, args);
            return result;
        }
    }
}