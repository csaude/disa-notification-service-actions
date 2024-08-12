package disa.notification.service.repository;

import java.util.List;

import org.springframework.data.jpa.repository.EntityGraph;
import org.springframework.data.jpa.repository.EntityGraph.EntityGraphType;
import org.springframework.data.repository.Repository;

import disa.notification.service.entity.ImplementingPartner;

public interface ImplementingPartnerRepository extends Repository<ImplementingPartner, Integer> {
    @EntityGraph(value = "graph.ImplementingPartner.orgUnits", type = EntityGraphType.LOAD)
    List<ImplementingPartner> findByEnabledTrue();
}
