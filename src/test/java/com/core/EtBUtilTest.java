package com.core;

import com.beta.Emp;
import org.junit.Test;

import java.util.List;

public class EtBUtilTest {

	@Test
	public void test1(){
		EtBUtil<Emp> tool = EtBUtil.<Emp>create(Emp.class);

		try {
			List<Emp> emps = tool.read(getClass().getResource("/").getPath() + "source.xls");
			for (Emp emp : emps) {
				System.out.println(emp);
			}
		} catch (IllegalAccessException e) {
			e.printStackTrace();
		} catch (InstantiationException e) {
			e.printStackTrace();
		}
	}
}
