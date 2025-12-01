package com.lizhiwei.quickExcel.core;

import cn.idev.excel.FastExcel;
import cn.idev.excel.support.ExcelTypeEnum;
import com.lizhiwei.quickExcel.entity.ExcelEntity;
import com.lizhiwei.quickExcel.entity.Since;
import com.lizhiwei.quickExcel.model.FileOperation;
import com.lizhiwei.quickExcel.model.RowModel;
import com.lizhiwei.quickExcel.model.SheetModel;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.BiConsumer;
import java.util.stream.IntStream;
import java.util.stream.Stream;

/**
 * FastExcel 抽象层，提供 QuickExcel 转换至 FastExcel
 */
public class FastExcelCore extends ExcelUtil {
    public static <T> Stream<List<T>> chunkedStream(List<T> list, int chunkSize) {
        return IntStream.iterate(0, i -> i + chunkSize)
                .limit((list.size() + chunkSize - 1) / chunkSize)
                .mapToObj(i -> list.subList(i, Math.min(i + chunkSize, list.size())));
    }

    public <T> void createExcel(FileOperation operation, Class<T> entity, List<T> listContent) {
        operation.run(outputStream -> {
            List<ExcelEntity> top = getExcelEntities(entity);
            var sheet = FastExcel.write(outputStream).excelType(ExcelTypeEnum.XLSX).sheet();
            AtomicInteger index = new AtomicInteger(0);
            List<List<String>> head = Collections.singletonList(top.stream().flatMap(x -> Stream.of(x.getTitle())).toList());
            sheet.table(index.get()).head(head);
            chunkedStream(listContent, 10000).forEach((x) -> {
                sheet.table(index.incrementAndGet()).doWrite(() -> {
                    List<Map<String, String>> list = new ArrayList<>();
                    for (T t : x) {
                        Map<String, String> map = new HashMap<>();
                        //获取类属性
                        for (ExcelEntity excelEntity : top) {
                            String value = null;
                            try {
                                value = getParamString(excelEntity, t);
                            } catch (NoSuchFieldException | IllegalAccessException e) {
                                throw new RuntimeException(e);
                            }
                            //循环设置每列的值
                            map.put(excelEntity.getTitle(), value);
                            break;
                        }
                        list.add(map);
                    }
                    return list;
                });
            });
        });


    }

    @Override
    public <T> SheetModel setSheetContent(SheetModel sheet, List<T> listContent, List<ExcelEntity> listTitle, List<Since> since, CellStyle cs, short ss, BiConsumer<T, RowModel> row) {
        return null;
    }
}
