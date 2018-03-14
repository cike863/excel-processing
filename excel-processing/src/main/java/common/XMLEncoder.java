package common;

import org.apache.commons.lang3.StringUtils;
import util.RegexUtil;

import java.util.regex.PatternSyntaxException;

/**
 * @version: v1.0
 * @author: zhulong
 * project: order
 * copyright: BLISSMALL TECHNOLOGY CO., LTD. (c) 2015-2016
 * createTime: 2017/3/1 0001 下午 4:50
 * modifyTime:
 * modifyBy:
 */
public class XMLEncoder {
    private static final String[] xmlCode = new String[256];

    static {
        // Special characters
        xmlCode['&'] = " "; // ampersand
        xmlCode['&'] = " "; // ampersand
        xmlCode['<'] = "{"; // lower than
        xmlCode['>'] = "}"; // greater than
        xmlCode['\n'] = " "; // greater than
        xmlCode['\u0007'] = " "; // greater than
        xmlCode['\''] = " "; // greater than
        xmlCode['\"'] = " "; // greater than
        //+++
        //
    }

    /**
     * <p>
     * Encode the given text into xml.
     * </p>
     *
     * @param string the text to encode
     * @return the encoded string
     */
    public static String encode(String string) {
        if (string == null){ return "";}
        int n = string.length();
        char character;
        String xmlchar;
        StringBuffer buffer = new StringBuffer();
        // loop over all the characters of the String.
        for (int i = 0; i < n; i++) {
            character = string.charAt(i);
            // the xmlcode of these characters are added to a StringBuffer one by one
            try {
                xmlchar = xmlCode[character];
                if (xmlchar == null) {
                    buffer.append(character);
                } else {
                    buffer.append(xmlCode[character]);
                }
            } catch (ArrayIndexOutOfBoundsException aioobe) {
                buffer.append(character);
            }
        }
        return buffer.toString();
    }

    public static String StringFilter(String text) throws PatternSyntaxException {
        // 只允许字母和数字 // String regEx ="[^a-zA-Z0-9]";
        // 清除掉所有特殊字符
        if (StringUtils.isBlank(text) && StringUtils.isBlank(text.trim())) {
            return text;
        }
        boolean check = RegexUtil.text(text);
        if (check) {
            return text;
        }
        text = text.trim();
        int size = text.length() - 1;
        String str = "";
        String temp = "";
        for (int i = 0; i < size; i++) {
            str = text.substring(i, i + 1);
            check = RegexUtil.text(str);
            if (check) {
                temp += str;
            }
        }
        return temp;
    }

    public static void main(String[] args) {
        System.out.println(XMLEncoder.encode("亲爱的老婆：\n" +
                "相亲相爱，鸿福满堂。\n" +
                "Love You~每天都多一点点……\n" +
                "                                小醉&Sugar"));
    }

}
