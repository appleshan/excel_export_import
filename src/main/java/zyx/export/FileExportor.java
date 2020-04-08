package zyx.export;

import org.apache.poi.ss.usermodel.Workbook;
import zyx.export.domain.common.ExportConfig;
import zyx.export.domain.common.ExportResult;
import zyx.export.domain.common.ExportType;
import zyx.export.domain.excel.ExportCSVResult;
import zyx.export.domain.excel.ExportExcelResult;
import zyx.export.exception.FileExportException;
import zyx.export.impl.CSVExportor;
import zyx.export.impl.Excel2003Exportor;
import zyx.export.impl.Excel2007Exportor;

import java.util.List;

/**
 * Created by stark.zhang on 2015/11/7.
 */
public class FileExportor {
    public final static String EXPORT_XML_BASE_PATH = "/properties/framework/export/";

    /**
     * 通过list<T> T可为bean或者map<String, Object>  导出文件
     *
     * @param exportConfig
     * @param data
     * @return
     * @throws zyx.export.exception.FileExportException
     */
    public static ExportResult getExportResult(ExportConfig exportConfig, List<?> data) throws FileExportException {
        ExportType exportType = exportConfig.getExportType();
        switch (exportType) {
            // index: 0
            case EXCEL2003:
                Workbook workbookOfExcel2003 = new Excel2003Exportor()
                        .getExportResult(data, exportConfig.getExportCells());

                ExportExcelResult exportExcel2003Result = new ExportExcelResult();
                exportExcel2003Result.setWorkbook(workbookOfExcel2003);
                exportExcel2003Result.setFileName(exportConfig.getFileName());

                return exportExcel2003Result;

            // index: 1
            case EXCEL2007:
                Workbook workbookOfExcel2007 = new Excel2007Exportor()
                        .getExportResult(data, exportConfig.getExportCells());

                ExportExcelResult exportExcel2007Result = new ExportExcelResult();
                exportExcel2007Result.setWorkbook(workbookOfExcel2007);
                exportExcel2007Result.setFileName(exportConfig.getFileName());

                return exportExcel2007Result;

            // index: 2
            case CSV:
                StringBuilder stringBuilder = new CSVExportor()
                        .getExportResult(data, exportConfig.getExportCells());

                ExportCSVResult exportCSVResult = new ExportCSVResult();
                exportCSVResult.setResult(stringBuilder.toString());
                exportCSVResult.setFileName(exportConfig.getFileName());

                return exportCSVResult;
        }

        throw new FileExportException("找不到对应的 export type, export type is " + exportType.getNumber());
    }
}
