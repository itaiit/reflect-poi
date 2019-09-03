package com.beta;

import java.io.Serializable;
import java.sql.Date;

/**
 * @author 22640
 */
public class Emp implements Serializable {
    /*
     * 实现对象在网络中传输或存到硬盘上，需要实现序列化接口
     */

    private static final long serialVersionUID = 1L;

    /*
     * Integer比基本数据类型多了一个null类型，
     * 而数据库中的值默认都是null，使用Integer类型操作较方便
     */
    private Integer empno;
    private String ename;
    private String job;
    private Integer mgr;
    private Date hiredate;
    private Double sal;
    private Double comm;
    private Integer deptno;
    private String newadd;

    public String getNewadd() {
        return newadd;
    }

    public void setNewadd(String newadd) {
        this.newadd = newadd;
    }

    public Emp() {
    }

    public Integer getEmpno() {
        return empno;
    }

    public void setEmpno(Integer empno) {
        this.empno = empno;
    }

    public String getEname() {
        return ename;
    }

    public void setEname(String ename) {
        this.ename = ename;
    }

    public String getJob() {
        return job;
    }

    public void setJob(String job) {
        this.job = job;
    }

    public Integer getMgr() {
        return mgr;
    }

    public void setMgr(Integer mgr) {
        this.mgr = mgr;
    }

    public Date getHiredate() {
        return hiredate;
    }

    public void setHiredate(Date hiredate) {
        this.hiredate = hiredate;
    }

    public Double getSal() {
        return sal;
    }

    public void setSal(Double sal) {
        this.sal = sal;
    }

    public Double getComm() {
        return comm;
    }

    public void setComm(Double comm) {
        this.comm = comm;
    }

    public Integer getDeptno() {
        return deptno;
    }

    public void setDeptno(Integer deptno) {
        this.deptno = deptno;
    }

    @Override
    public String toString() {
        return "Emp{" +
                "empno=" + empno +
                ", ename='" + ename + '\'' +
                ", job='" + job + '\'' +
                ", mgr=" + mgr +
                ", hiredate=" + hiredate +
                ", sal=" + sal +
                ", comm=" + comm +
                ", deptno=" + deptno +
                ", newadd='" + newadd + '\'' +
                '}';
    }
}
