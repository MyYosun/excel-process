package com.ling;

public class Person {
    private String id;      // 编号
    private String name;    // 姓名
    private Integer status; // 出场状态  0 -- 已出场(可入场) / 1 -- 未出场(不可入场)

    public String getId() {
        return id;
    }

    public void setId(String id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public Integer getStatus() {
        return status;
    }

    public void setStatus(Integer status) {
        this.status = status;
    }
}