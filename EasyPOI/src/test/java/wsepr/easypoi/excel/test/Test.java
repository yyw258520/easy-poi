package wsepr.easypoi.excel.test;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

public class Test {

	/**
	 * @param args
	 * @throws NoSuchMethodException 
	 * @throws SecurityException 
	 */
	public static void main(String[] args){
		Test test = new Test();
		
		try {
			Method say = Test.class.getMethod("say", String.class, int.class);
			long start = System.currentTimeMillis();
			for(int i=0;i<1000000;i++){
				say.invoke(test,"hello", 5);
			}
			long end = System.currentTimeMillis();
			System.out.println(end - start);
			start = System.currentTimeMillis();
			for(int i=0;i<1000000;i++){
				test.say("hello",5);
			}
			end = System.currentTimeMillis();
			System.out.println(end - start);
		} catch (IllegalArgumentException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IllegalAccessException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (InvocationTargetException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (SecurityException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (NoSuchMethodException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}
	
	public void say(String hello, int n){
		for(int i=0;i<n;i++){
			n+= i;
		}
	}

}
