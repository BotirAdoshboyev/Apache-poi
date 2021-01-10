package com.simple.apache

import org.springframework.boot.autoconfigure.SpringBootApplication
import org.springframework.boot.runApplication

@SpringBootApplication
class ApacheApplication

fun main(args: Array<String>) {
    runApplication<ApacheApplication>(*args)
}