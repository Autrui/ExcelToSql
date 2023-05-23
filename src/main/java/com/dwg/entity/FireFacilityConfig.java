package com.dwg.entity;

import cn.hutool.json.JSONObject;
import lombok.Data;

import java.util.List;

/**
 * @author Autrui
 * @date 2023/5/23
 * @apiNote
 */
@Data
public class FireFacilityConfig {

    String name;

    String code;

    Integer type;

    Integer del_flag;

    List<JSONObject> maintenance_config;

}
