package com.datn.application.repository;

import com.datn.application.entity.Statistic;
import com.datn.application.model.dto.StatisticDTO;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.data.jpa.repository.Query;
import org.springframework.stereotype.Repository;

import java.util.List;

@Repository
public interface StatisticRepository extends JpaRepository<Statistic, Long> {

    @Query(name = "getStatistic30Day",nativeQuery = true)
    List<StatisticDTO> getStatistic30Day();

    @Query(name = "getStatisticDayByDay",nativeQuery = true)
    List<StatisticDTO> getStatisticDayByDay(String toDate, String formDate);

    @Query(value = "SELECT * FROM statistic  WHERE date_format(created_at,'%Y-%m-%d') = date_format(NOW(),'%Y-%m-%d')",nativeQuery = true)
    Statistic findByCreatedAT();

    @Query(value= "SELECT sum(sales) FROM statistic;", nativeQuery=true)
    long calcRevenues();
    @Query(value= "SELECT sum(profit) FROM statistic;", nativeQuery=true)
    long calcProfits();
}
