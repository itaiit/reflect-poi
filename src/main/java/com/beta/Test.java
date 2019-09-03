
package com.beta;

/**
 * 
 * @author itaiit
 *
 * @param <T>
 */
public class Test<T> {

	@SuppressWarnings("unused")
	private Class<T> cla = null;
	
	public Test(Class<T> cla) {
		this.cla = cla;
	}
	
	public static <T> Test<T> getObj(Class<T> type){
		return new Test<T>(type);
	}
	
}
