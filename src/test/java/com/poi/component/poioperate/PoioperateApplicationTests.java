package com.poi.component.poioperate;

import com.alibaba.fastjson.JSONArray;
import com.poi.component.poioperate.utils.ExcelReader;
import com.poi.component.poioperate.utils.ExcelWriter;
import com.poi.component.poioperate.vo.DbTableIncrement;
import com.poi.component.poioperate.vo.DbTableResultVO;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Test;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.util.StringUtils;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.*;
import java.util.stream.Collectors;

@SpringBootTest
class PoioperateApplicationTests {

	@Test
	void contextLoads() {
		String t = "\u202C23225";
		System.out.println(t.substring(0, 1));
		String increment = "1978";
		if (increment.length() <= 2) {
			increment = "200";
		} else {
			String incr = increment.substring(0, 1);
			int i = Integer.valueOf(incr) + 3;
			int length = increment.length();
			increment = String.valueOf(i);
			for (int j = 1; j < length; j++) {
				increment+="0";
			}
		}
		System.out.println(increment);
	}

	@Test
	void readExcel(){
		List<DbTableIncrement> agencyList = ExcelReader.readExcel("C:\\Users\\lenovo\\Desktop\\工作簿1.xlsx", DbTableIncrement.class);
		List<DbTableIncrement> ydzList = ExcelReader.readExcel("C:\\Users\\lenovo\\Desktop\\数据表情况汇总.xlsx", DbTableIncrement.class);
		Map<String, DbTableResultVO> map = new HashMap<>();
		for (DbTableIncrement ydzVO : ydzList) {
			DbTableResultVO dbVO = new DbTableResultVO();
			dbVO.setYdzTableName(ydzVO.getTableName());
			dbVO.setYdzIncr(ydzVO.getIncrement());
			map.put(ydzVO.getTableName(), dbVO);
		}
		for (int k = 0; k < agencyList.size(); k++) {
			DbTableIncrement agencyVO = agencyList.get(k);
			String tableName = "agency_" + agencyVO.getTableName();
			if (map.containsKey(tableName)) {
				DbTableResultVO dbVO = map.get(tableName);
				dbVO.setAgencyTableName(agencyVO.getTableName());
				dbVO.setAgencyIncr(agencyVO.getIncrement());
				String increment = agencyVO.getIncrement();
				System.out.println("i = " + k + "str = '" + increment + "'");
				if (increment.length() <= 2) {
					increment = "200";
				} else {
					String incr = increment.substring(0, 1);
					System.out.println("incr = " + incr);
					int i = Integer.valueOf(incr) + 3;
					int length = increment.length();
					increment = String.valueOf(i);
					for (int j = 1; j < length; j++) {
						increment+="0";
					}
				}
				dbVO.setUpdateIncr(increment);
			}
		}
		List<DbTableResultVO> tableList = map.entrySet().parallelStream().map(Map.Entry::getValue).
				filter(a -> StringUtils.isEmpty(a.getAgencyTableName())).collect(Collectors.toList());
		List<DbTableResultVO> result = map.entrySet().parallelStream().map(Map.Entry::getValue)
				.filter(a-> !StringUtils.isEmpty(a.getAgencyTableName())).collect(Collectors.toList());
		try {
			FileOutputStream fos = new FileOutputStream("C:\\Users\\lenovo\\Desktop\\new.xlsx");
			Workbook workbook = ExcelWriter.exportData(result);
			workbook.write(fos);
			fos.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
		StringBuffer strBuffer = new StringBuffer();
		for (int i = 0; i < result.size(); i++) {
			strBuffer.append("ALTER TABLE ");
			strBuffer.append(result.get(i).getYdzTableName());
			strBuffer.append(" AUTO_INCREMENT = ");
			strBuffer.append(result.get(i).getUpdateIncr());
			strBuffer.append(";\n");
		}
		try {
			FileOutputStream fos = new FileOutputStream("C:\\Users\\lenovo\\Desktop\\execute.sql");
			fos.write(strBuffer.toString().getBytes());
			fos.close();
		}catch (Exception e) {
			e.printStackTrace();
		}
		System.out.println(tableList.parallelStream().map(DbTableResultVO::getYdzTableName).collect(Collectors.toList()));
	}

}
