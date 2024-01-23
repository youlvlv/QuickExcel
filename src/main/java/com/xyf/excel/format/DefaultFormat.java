package com.xyf.excel.format;

import com.xyf.excel.entity.ClassMap;
import org.reflections.Reflections;
import org.reflections.util.ConfigurationBuilder;

import java.lang.reflect.InvocationTargetException;
import java.util.Set;

/**
 * 默认转换器，默认扫描format.type包下所有的类型转换器
 * 可以使用ExcelConfig加载类型转换器
 *
 * @author lizhiwei
 */
public class DefaultFormat implements ExcelFormat<Object> {

    public static final ClassMap CLASS_FORMAT_MAP = new ClassMap();

    static {
        Reflections reflections = new Reflections(new ConfigurationBuilder().forPackages("com.lizhiwei.quickExcel.format.type"));
        Set<Class<? extends ExcelFormatByType>> classes = reflections.getSubTypesOf(ExcelFormatByType.class);
        classes.forEach(x -> {
            ExcelFormatByType<?> typeFormat;
            try {
                typeFormat = x.getDeclaredConstructor().newInstance();
            } catch (InstantiationException | IllegalAccessException | InvocationTargetException |
                     NoSuchMethodException e) {
                throw new RuntimeException(e);
            }

            CLASS_FORMAT_MAP.put(typeFormat.getType(), typeFormat);
        });
    }

//    public static final DateTimeFormatter fmt = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");


    @Override
    public String WriterToExcel(Object v) {
        if (v != null) {
            return CLASS_FORMAT_MAP.get(v.getClass()).writerToExcel(v);
        } else {
            return "";
        }
//        if (v != null) {
//
//            //判断属性的类型
//            if (v instanceof String) {
//                //String类型执行toString方法
//                return v.toString();
//            } else if (v instanceof Date) {
//                //时间类型，则转换时间
//                DateFormat dFormat = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss"); //HH表示24小时制；
//                return dFormat.format((Date) v);
//            } else if (v instanceof Number) {
//                return v.toString();
//            } else if (v instanceof Boolean) {
//                if ((boolean) v) {
//                    return "是";
//                } else {
//                    return "否";
//                }
//            } else if (v instanceof LocalDateTime) {
//                return ((LocalDateTime) v).format(fmt);
//            } else return v.toString();
//        }
//        return "";
    }

    @Override
    public Object ReadToExcel(String v) {
        return null;
    }

    public Object ReadToExcel(Class<?> type, String v) {
//        SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
        if (v == null || v.equals("")) {
            return null;
        }
        return CLASS_FORMAT_MAP.get(type).ReadToExcel(v);
//        if (v == null || v.equals("")) {
//            return null;
//        } else {
//            if (type == String.class) {
//                return v;
//            } else if (type == Integer.class) {
//                return Integer.valueOf(v);
//            } else if (type == Long.class) {
//                return Long.valueOf(v);
//            } else if (type == Double.class) {
//                return Double.valueOf(v);
//            } else if (type == Boolean.class) {
//                if (v.contains("是")) {
//                    return true;
//                } else if (v.contains("否")) {
//                    return false;
//                }
//                return Boolean.valueOf(v);
//            } else if (type == Date.class) {
//                try {
//                    return sdf.parse(v);
//                } catch (ParseException e) {
//                    throw new RuntimeException(e);
//                }
//            }
//        }
//        return null;
    }
}
