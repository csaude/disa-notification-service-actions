package disa.notification.service.entity;

import javax.persistence.Entity;
import javax.persistence.Id;

@Entity
public class OrgUnit {
    @Id
    private String code;

    private String province;

    private String district;

    private String facility;

    private Integer implementingPartnerId;

    public String getCode() {
        return code;
    }

    public void setCode(String code) {
        this.code = code;
    }

    public String getProvince() {
        return province;
    }

    public void setProvince(String province) {
        this.province = province;
    }

    public String getDistrict() {
        return district;
    }

    public void setDistrict(String district) {
        this.district = district;
    }

    public String getFacility() {
        return facility;
    }

    public void setFacility(String facility) {
        this.facility = facility;
    }

    public Integer getImplementingPartnerId() {
        return implementingPartnerId;
    }

    public void setImplementingPartnerId(Integer implementingPartnerId) {
        this.implementingPartnerId = implementingPartnerId;
    }

    @Override
    public String toString() {
        return "OrgUnit [code=" + code + ", facility=" + facility + "]";
    }
}
