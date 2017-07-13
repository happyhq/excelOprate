package test;

import java.io.File;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class WriteToExcel {
	public static void main(String[] args) {
		List<Map<String, Object>> list = new ArrayList<Map<String, Object>>();
		Map<String, Object> map;
		for (int i = 1; i < 10; i++) {
			map = new HashMap<String, Object>();
			map.put("id", i);
			map.put("age", 18);
			map.put("name", "张"+i);
			list.add(map);
		}
		exportExcel("D://a.xls", list);

	}

	/**
	 * 导出excel表格。 步骤：
	 * 
	 * 1.创建WritableWorkbook对象
	 * 
	 * 2.生成工作表
	 * 
	 * 3.将数据遍历获取，通过sheet.addCell(new Label(j, i+1, String.valueOf(map
	 * .get(info[j]))))方法将map .get(info[j])值保存到i+1行j列
	 * 
	 * @param path
	 * @param list
	 */
	@SuppressWarnings("unchecked")
	private static <T> void exportExcel(String path, List<T> list) {
		WritableWorkbook book = null;
		System.out.println(path);
		String info[] = new String[list.size()];
		try {
			// 创建WritableWorkbook对象
			book = Workbook.createWorkbook(new File(path));
			// 生成名为eccif的工作表，参数0表示第一页
			WritableSheet sheet = book.createSheet("sheet", 0);
			Map<String, Object> map = (Map<String, Object>) list.get(0);
			Iterator<Map.Entry<String, Object>> iterator = map.entrySet()
					.iterator();
			int i = 0;
			while (iterator.hasNext()) {
				Map.Entry<String, Object> entry = iterator.next();
				info[i++] = entry.getKey();
			}
			// 表头导航
			for (i = 0; i < info.length; i++) {
				//设置表头（第一行）
				Label label = new Label(i, 0, info[i]);
				sheet.addCell(label);
				map = (Map<String, Object>) list.get(i);
				//设置表格内容（第二行起）
				for (int j = 0; j < map.size(); j++) {
					sheet.addCell(new Label(j, i+1, String.valueOf(map
							.get(info[j]))));
				}
			}
			// 写入数据并关闭文件
			book.write();
		} catch (Exception e) {
			System.out.println(e);
		} finally {
			if (book != null) {
				try {
					book.close();
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		}
	}
}
