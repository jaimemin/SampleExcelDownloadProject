package com.tistory.jaimemin.excel.listener;

import com.tistory.jaimemin.excel.constant.ExcelConstants;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.lang3.ObjectUtils;

import javax.servlet.http.HttpSession;
import javax.servlet.http.HttpSessionEvent;
import javax.servlet.http.HttpSessionListener;
import java.io.File;
import java.util.HashSet;
import java.util.Set;

@Slf4j
public class SessionListener implements HttpSessionListener {

    private static final int SESSION_TIME = 1800;

    @Override
    public void sessionCreated(HttpSessionEvent se) {
        se.getSession().setMaxInactiveInterval(SESSION_TIME);
    }

    @Override
    public void sessionDestroyed(HttpSessionEvent se) {
        HttpSession session = se.getSession();
        Set<String> excelFilePathes = (HashSet<String>) session.getAttribute(ExcelConstants.SESSION_KEY);

        if (ObjectUtils.isEmpty(excelFilePathes)) {
            return;
        }

        for (String excelFilePath : excelFilePathes) {
            File fileToDelete = new File(excelFilePath);

            if (!fileToDelete.exists()) {
                continue;
            }

            deleteFile(fileToDelete);
        }
    }

    private synchronized void deleteFile(File file) {
        file.delete();
    }
}
