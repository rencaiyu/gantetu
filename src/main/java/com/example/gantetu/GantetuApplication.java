package com.example.gantetu;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * 应用启动入口。
 * <p>
 * 该类使用 {@link SpringBootApplication} 开启 Spring Boot 自动配置、组件扫描
 * 以及配置类能力，运行 main 方法后会拉起内嵌 Web 容器。
 */
@SpringBootApplication
public class GantetuApplication {

    /**
     * Java 程序主方法，负责启动整个 Spring Boot 应用上下文。
     *
     * @param args 启动参数
     */
    public static void main(String[] args) {
        SpringApplication.run(GantetuApplication.class, args);
    }
}
