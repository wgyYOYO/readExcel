package wgy.entity;

import lombok.Data;

@Data
public class User {
    String name;//姓名
    String sex;//性别
    String stuId;//学号
    String department;//院系
    String perfessional;//专业
    String birthYear;//出生年
    String birthMonth;//出生月
    String birthDay;//出生日
    String validityYear;//有效年
    String domitory;//宿舍
}
