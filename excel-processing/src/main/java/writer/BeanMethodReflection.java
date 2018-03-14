package writer;

import org.springframework.util.ConcurrentReferenceHashMap;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Map;
import java.util.Set;

/**
 * @Description: BeanGetMethod获取类
 * @author lizhang
 * @date 2017年3月29日 下午7:12:10
 */
public class BeanMethodReflection {

	private static final Map<Class<?>, Map<String, Method>> getMethodMap = new ConcurrentReferenceHashMap<>();

	private static final char[] get = { 'g', 'e', 't' };

	public static Map<String, Method> initAndGetReflection(Class<?> clazz) {
		Map<String, Method> getMethod = getMethodMap.get(clazz);
		if (getMethod == null) {
			getMethod = new HashMap<>();
			Set<String> fieldNameSet = new HashSet<>();
			for (Field field : clazz.getDeclaredFields()) {
				String name = field.getName();
				StringBuilder fieldName = new StringBuilder(name);
				char c = fieldName.charAt(0);
				if (c >= 'A' && c <= 'Z') {
					c = (char) (c + 32);
					fieldName.setCharAt(0, c);
				}
				fieldNameSet.add(fieldName.toString());
			}
			for (Method method : clazz.getDeclaredMethods()) {
				if (method.getParameterCount() == 0) {
					String methodName = method.getName();
					StringBuilder preFieldName = new StringBuilder(methodName.length());
					for (int i = 0; i < methodName.length(); i++) {
						char c = methodName.charAt(i);
						if (i < 3) {
							if (c != get[i]) {
								break;
							}
						} else {
							if (i == 3) {
								if (c >= 'a' && c <= 'z') {
									break;
								}
								if (c >= 'A' && c <= 'Z') {
									c = (char) (c + 32);
								}
							}
							preFieldName.append(c);
						}
					}
					String fieldName = preFieldName.toString();
					if (fieldNameSet.contains(fieldName)) {
						getMethod.put(fieldName, method);
					}
				}
			}
			getMethodMap.put(clazz, getMethod);
		}
		return getMethod;
	}

}
