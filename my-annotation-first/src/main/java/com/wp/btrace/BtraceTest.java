package com.wp.btrace;

/* BTrace Script Template */
import com.sun.btrace.annotations.*;
import static com.sun.btrace.BTraceUtils.*;

@BTrace
public class BtraceTest {
    /* put your code here */
    @OnMethod(clazz="com.alibaba.olapengine.service.impl.sync.OlapengineMessageReceiver",
            method="receive",
            location=@Location(Kind.RETURN)
    )
    public static void func() {
        println("调用堆栈！");
        jstack();

    }
}
