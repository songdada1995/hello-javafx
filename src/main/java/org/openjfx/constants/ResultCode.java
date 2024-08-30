package org.openjfx.constants;

/**
 * @author pcq
 * @date 2017/7/4
 */
public interface ResultCode {

    /**
     * 1001xxx
     */
    enum Common {
        //未知错误
        UNKNOWN_ERROR(1001000),
        //资源不存在
        NOT_FOUND(1001001),
        //参数非法
        ILLEGAL_ARGUMENT(1001002),
        //服务器内部错误
        INTERNAL_SERVER_ERROR(1001500);
        public int code;

        Common(int code) {
            this.code = code;
        }
    }

}
