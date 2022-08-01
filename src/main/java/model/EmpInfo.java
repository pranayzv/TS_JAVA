package model;

public class EmpInfo {
    private String filename;
    private int month;
    private int year;
    private String empName;
    private String approverName;
    private String loc;
    private String client;

    public EmpInfo(){}

    public EmpInfo(String filename, int month, int year, String empName, String approverName, String loc, String client) {
        this.filename = filename;
        this.month = month;
        this.year = year;
        this.empName = empName;
        this.approverName = approverName;
        this.loc = loc;
        this.client = client;
    }

    public String getFilename() {
        return filename;
    }

    public void setFilename(String filename) {
        this.filename = filename;
    }

    public int getMonth() {
        return month;
    }

    public void setMonth(int month) {
        this.month = month;
    }

    public int getYear() {
        return year;
    }

    public void setYear(int year) {
        this.year = year;
    }

    public String getEmpName() {
        return empName;
    }

    public void setEmpName(String empName) {
        this.empName = empName;
    }

    public String getApproverName() {
        return approverName;
    }

    public void setApproverName(String approverName) {
        this.approverName = approverName;
    }

    public String getLoc() {
        return loc;
    }

    public void setLoc(String loc) {
        this.loc = loc;
    }

    public String getClient() {
        return client;
    }

    public void setClient(String client) {
        this.client = client;
    }

    @Override
    public String toString() {
        return "model.EmpInfo{" +
                "filename='" + filename + '\'' +
                ", month=" + month +
                ", year=" + year +
                ", empName='" + empName + '\'' +
                ", approverName='" + approverName + '\'' +
                ", loc='" + loc + '\'' +
                ", client='" + client + '\'' +
                '}';
    }
}

