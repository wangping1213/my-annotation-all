package com.wp.btrace;

/* BTrace Script Template */
import com.sun.btrace.annotations.*;
import static com.sun.btrace.BTraceUtils.*;

@BTrace
public class BtraceForTime {

    @TLS private static long startTime = 0;

    @OnMethod(clazz="com.wp.mvc.web.AAA",
            method="getPerson"
    )
    public static void startAAA() {
        startTime = timeNanos();
    }

    @OnMethod(clazz="com.wp.mvc.web.AAA",
            method="getPerson",
            location=@Location(Kind.RETURN)
    )
    public static void endAAA(@Duration long duration, @Return Object obj) {
        long time = timeNanos() - startTime;
        println(strcat("execute time(nanos): ", str(time)));
        println(strcat("duration(nanos): ", str(duration)));
        println(str(obj));
        jstack();
    }
}
