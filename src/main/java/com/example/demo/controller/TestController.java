package com.example.demo.controller;

import com.example.demo.model.User;
import com.example.demo.util.ExcelUtils;
import com.example.demo.util.ReadRowMapper;
import com.example.demo.util.WriteRowMapper;
import org.apache.poi.ss.usermodel.Row;
import org.springframework.stereotype.Controller;
import org.springframework.ui.Model;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.ResponseBody;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

/**
 * @author lihuasheng
 * @since 2019/4/29 21:25
 */
@Controller
@RequestMapping("/")
public class TestController {

    public static final String[] TITLE = {"姓名","年龄","账号","密码"};

    @RequestMapping("/")
    public String index(Model model) {
        return "index";
    }

    @RequestMapping(value = "/importExcel")
    @ResponseBody
    public void importExcel(@RequestParam("file")MultipartFile file) throws IOException {
        List<User> list = ExcelUtils.importExcel(file, (row, map) -> {
            User user = new User();
            user.setName((String) map.get("姓名"));
            user.setAge(String.valueOf(map.get("年龄")));
            user.setAccount((String) map.get("账号"));
            user.setPassword(String.valueOf(map.get("密码")));
            return user;
        });
        list.forEach(System.out::println);
    }

    @RequestMapping(value = "/exportExcel")
    @ResponseBody
    public void exportExcel(HttpServletResponse response) throws IOException {
        List<User> list = new ArrayList<>();
        User user = new User();
        user.setName("逗比");
        user.setAge("1");
        user.setAccount("121313");
        user.setPassword("rfgh");
        list.add(user);
        ExcelUtils.exportExcel(response, "前方高能", "sheet", "测试用户", TITLE, list, param -> {
            List<String> value = new ArrayList<>(TITLE.length);
            User user1 = (User) param;
            value.add(user1.getName());
            value.add(user1.getAge());
            value.add(user1.getAccount());
            value.add(user1.getPassword());
            return value;
        });
    }


}