package com.example.gantetu;

import org.mybatis.spring.annotation.MapperScan;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
@MapperScan("com.example.gantetu.mapper")
public class GantetuApplication {

    /**
     * Spring Boot 应用入口，同时启用 MyBatis Mapper 扫描。
     */
    public static void main(String[] args) {
        SpringApplication.run(GantetuApplication.class, args);
    }
}
