package controller;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import br.com.equipoinfo.hoc.business.exception.BusinessException;
import br.com.equipoinfo.hoc.dao.pojo.FiltroRelatorio;

public class Teste {

	
	public ByteArrayOutputStream relatorioExcel(
			FiltroRelatorio filtroRelatorio) throws BusinessException {
		List<Object> listaComAsInfomacoes = null;
		ByteArrayOutputStream retorno = null;
		FileInputStream template = null;
		XSSFWorkbook excel = null;
		try {
			//nome da pasta do arquivo
			String pathTemplates = "C:/apachepoi/templantes/excel";

			//nome do arquivo
			template = new FileInputStream(pathTemplates+ "/relatorio_exemplo.xlsx");
			excel = new XSSFWorkbook(template);
			
			//dao que busca as info no banco
			listaComAsInfomacoes = relatorioDAO.getRelatorioExemplo(filtroRelatorio);
			if (listaComAsInfomacoes == null || listaComAsInfomacoes.size() == 0) {
				throw new BusinessException(
						"Não foram encontrados registros para o filtro informado!");

			} else {
				//criando primeira planilha
				Sheet sheet = excel.getSheetAt(0);
				Row row;
				
				//tratamento especial, para as datas já virem formatadas no formato Data do Excel
				CellStyle cellStyle = excel.createCellStyle();
				CreationHelper createHelper = excel.getCreationHelper();
				cellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd/MM/yyyy"));
				int i = 2;

				for (Object doc : listaComAsInfomacoes) {
					Object[] aux = (Object[]) doc;
					row = sheet.createRow((int) i);

					// cliente
					row.createCell(0).setCellValue(aux[0] == null ? " " : aux[0].toString());
					// cnpj
					row.createCell(1).setCellValue(aux[1] == null ? " " : aux[1].toString());
					// razao social
					row.createCell(2).setCellValue(aux[2] == null ? " " : aux[2].toString());
					// numero
					row.createCell(3).setCellValue(aux[3] == null ? " " : aux[3].toString());
					
					// data 
					if (aux[4] == null)
						row.createCell(4).setCellValue(" ");
					else {
						Cell cell = row.createCell(4);
						cell.setCellValue((Date) aux[4]);
						cell.setCellStyle(cellStyle);
					}
					// data 
					if (aux[5] == null)
						row.createCell(5).setCellValue(" ");
					else {
						Cell cell = row.createCell(5);
						cell.setCellValue((Date) aux[5]);
						cell.setCellStyle(cellStyle);
					}
					// valor
					row.createCell(6).setCellValue((double) (aux[6] == null ? " " : aux[6]));	
					i++;
					
					
					//criando segunda planilha dentro do mesmo arquivo
					Sheet segundaPlanilha = excel.getSheetAt(1);
					Row linhaSegundaPlanilha;
					
					//alimentando a segunda planilha
					linhaSegundaPlanilha = segundaPlanilha.createRow((int) 2);

					//volume total do(s) dia(s)
					linhaSegundaPlanilha.createCell(0).setCellValue(listaComAsInfomacoes.size());				
					
					// valor total do(s) dia(s) 
					String strFormula= "SUM(Plan1!I3:I1048576)";
					Cell cell = row.createCell(1);
					cell.setCellType(CellType.FORMULA);
					linhaSegundaPlanilha.createCell(1).setCellFormula(strFormula);
							
					//valor do ticket médio (valor total do dia / volume total do dia
					linhaSegundaPlanilha.createCell(2).setCellFormula("B3/A3");
				}
				retorno = new ByteArrayOutputStream();
				excel.write(retorno);
				template.close();
			}
		} catch (Exception e) {
			throw new BusinessException(e.getMessage());
		}
		return retorno;
	}
}
