package com.example.poi.controller;

import com.example.poi.entity.User;
import com.example.poi.util.ExportExcelUtil;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import java.io.File;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.*;

/**
 * Created by wangliang on 2017/8/31.
 */
@Controller
@RequestMapping("/users")
public class UserController {

    @GetMapping("/v1/list")
    public String index(HttpSession session) {
        List<User> list = new ArrayList<>();
        User user1 = new User("王亮", "26", 1, "727474430@qq.com", "18215518314");
        User user2 = new User("吕佳", "29", 0, "807724703@qq.com", "18202847188");
        User user3 = new User("小兔兔", "18", 0, "tutu@163.com", "18688888888");
        User user4 = new User("小龟龟", "20", 1, "guigui@163.com", "18699999999");
        list.add(user1);
        list.add(user2);
        list.add(user3);
        list.add(user4);
        session.setAttribute("users", list);
        return "user";
    }

    /**
     * Export Excel Api
     *
     * @param request
     * @param response
     * @return
     */
    @PostMapping("/v1/export")
    public String exportExcel(HttpServletRequest request, HttpServletResponse response) {
        ServletOutputStream out = null;
        try {
            List<User> list = (List<User>) request.getSession().getAttribute("users");
            Map<String, String> titleMap = new HashMap<>();
            titleMap.put("name", "姓名");
            titleMap.put("age", "年龄");
            titleMap.put("sex", "性别");
            titleMap.put("email", "邮箱");
            titleMap.put("phone", "手机");
            if (list != null) {
                response.setContentType("application/vnd.ms-excel;charset=gb2312");
                response.setHeader("Content-Disposition", "attachment;filename = " + geneFileName());
                out = response.getOutputStream();
                ExportExcelUtil.exportExcel(list, "用户列表", titleMap, out);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    /**
     * Import Excel Api
     *
     * @param request
     * @param response
     * @param file
     * @return
     */
    @PostMapping("/v1/import")
    public String importExcel(HttpServletRequest request, HttpServletResponse response, MultipartFile file) {
        if (file != null && !file.isEmpty()) {
            String filePath = request.getSession().getServletContext().getRealPath("/") + file.getOriginalFilename();
            try {
                file.transferTo(new File(filePath));
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
        return "redirect:/users/v1/upload";
    }

    /**
     * Find Specify File And Import
     *
     * @param request
     * @param response
     * @return
     */
    @RequestMapping("/v1/upload")
    public String fileUpload(HttpServletRequest request, HttpServletResponse response) {
        String filePath = request.getSession().getServletContext().getRealPath("/");
        File uploadDest = new File(filePath);
        String[] fileNames = uploadDest.list();
        for (int i = 0; i < fileNames.length; i++) {
            //打印出文件名
            System.out.println(fileNames[i]);
            List<User> list = ExportExcelUtil.importExcel(new File(filePath + fileNames[i]));
            List<User> old = (List<User>) request.getSession().getAttribute("users");
            old.addAll(list);
        }
        return "user";
    }

    public String geneFileName() {
        return new SimpleDateFormat("yyyyMMddHHmmssSSS").format(new Date()) + ".xls";
    }
}
