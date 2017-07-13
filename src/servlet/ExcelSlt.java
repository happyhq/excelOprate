package servlet;

import java.io.IOException;
import java.util.Iterator;
import java.util.List;

import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import jxl.*;
import org.apache.commons.fileupload.*;

public class ExcelSlt extends HttpServlet {
	private static final long serialVersionUID = 1L;

	public void doGet(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		doPost(request, response);
	}

	public void doPost(HttpServletRequest request, HttpServletResponse response)
			throws ServletException, IOException {
		String tempPath = "";
		// 代表一个EXCEL文件
		Workbook wb = null;
		try {
			DiskFileUpload fu = new DiskFileUpload(); // 设置最大文件尺寸，这里是4MB
			fu.setSizeMax(4194304); // 设置缓冲区大小，这里是4kb
			fu.setSizeThreshold(4096); // 设置临时目录：
			fu.setRepositoryPath(tempPath); // 得到所有的文件：
			List fileItems = fu.parseRequest(request);
			Iterator i = fileItems.iterator(); // 依次处理每一个文件：
			while (i.hasNext()) {
				FileItem fi = (FileItem) i.next(); // 获得文件名，这个文件名包括路径：
				String fileName = fi.getName(); // 在这里可以记录用户和文件信息
				wb = Workbook.getWorkbook(fi.getInputStream());
				if (wb == null) {
					return;
				}
				// 得到excel 所有工作表
				Sheet[] sheets = wb.getSheets();
				if (sheets != null) {
					for (int c = 0; c < sheets.length; c++) {
						// 遍历各个工作表
						Sheet s = sheets[c];
						int columns = s.getColumns();
						int rows = s.getRows();
						System.out.println("sheet[" + c + "]  columns:"
								+ columns + "  rows:" + rows);
						if (columns > 0 || rows > 0) {
							for (int r = 0; r < rows; r++) {
								for (int col = 0; col < columns; col++) {
									// 单元格getCell (行，列)
									Cell cell = s.getCell(col, r);
									System.out.print(cell.getContents() + "  ");// 输出单元格数据
								}
								System.out.println();
							}
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (wb != null) {
				wb.close();
			}
		}
	}

}
