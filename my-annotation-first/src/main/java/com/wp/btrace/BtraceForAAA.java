package com.wp.btrace;
import com.sun.btrace.annotations.*;

import static com.sun.btrace.BTraceUtils.*;

@BTrace
public class BtraceForAAA {
    /* put your code here */
    @OnMethod(clazz = "com.wp.mvc.web.AAA",
            method = "aaa",
            location = @Location(Kind.RETURN)
    )
    public static void func() {
        println("begine!");
        jstack();
    }
}