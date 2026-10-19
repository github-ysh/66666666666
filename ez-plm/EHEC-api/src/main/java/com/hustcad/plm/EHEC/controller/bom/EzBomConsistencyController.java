/*
 * Copyright © 2021 天喻软件
 *
 * Licensed under the Apache License, Version 2.0 (the "License"); you may not use this file except in compliance with
 * the License. You may obtain a copy of the License at
 *
 * http://www.apache.org/licenses/LICENSE-2.0
 *
 * Unless required by applicable law or agreed to in writing, software distributed under the License is distributed on
 * an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied. See the License for the
 * specific language governing permissions and limitations under the License.
 */
package com.hustcad.plm.EHEC.controller.bom;

import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.write.metadata.style.WriteCellStyle;
import com.alibaba.excel.write.metadata.style.WriteFont;
import com.alibaba.excel.write.style.HorizontalCellStyleStrategy;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import com.hustcad.plm.EHEC.bom.service.EzBomConsistencyService;
import com.hustcad.plm.EHEC.dto.EzCheckInSyncResultVO;
import com.ty.basic.response.ResponseResult;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.springframework.http.MediaType;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import javax.servlet.http.HttpServletResponse;
import org.springframework.util.StringUtils;
import javax.annotation.Resource;
import java.io.IOException;
import java.math.BigInteger;
import java.net.URLEncoder;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;

/**
 * @author TT
 * @className EzBomConsistencyController
 * @date 2026/3/19 11:43
 */
@Slf4j
@RestController
@Api(tags = {"二重-bom借用(备件转换)"})
@RequestMapping({"/rest/v1/ehec/bom/consistency"})
public class EzBomConsistencyController {
    @Resource
    private EzBomConsistencyService ezBomConsistencyService;

    @ApiOperation(value = "传入一个订货信息(项)节点的oid,自动同步对应的种( 图号件/非图号件)的下级结构", notes = "传入一个订货信息(项)节点的oid,自动同步对应的种( 图号件/非图号件)的下级结构", produces = "application/json; charset=utf-8")
    @GetMapping({"/syncXtoZConsistency"})
    public void syncXtoZConsistency(BigInteger partOid, String rkdyMasterOid) {
        ezBomConsistencyService.syncXtoZConsistency(partOid, rkdyMasterOid);
    }

    @ApiOperation(value = "根据种的bom结构创建项的bom结构", notes = "根据种的bom结构创建项的bom结构", produces = "application/json; charset=utf-8")
    @GetMapping({"/syncZtoXConsistency"})
    public void syncZtoXConsistency(BigInteger partOid, String partNo) {
        ezBomConsistencyService.syncZtoXConsistency(partOid, partNo);
    }

    @ApiOperation(value = "根据种的bom结构只同步指定的项的下级结构", notes = "传入种oid、入库单元编号和指定项OID列表(逗号分隔),只同步指定的项。targetXiangOids为空则同步全部", produces = "application/json; charset=utf-8")
    @GetMapping({"/syncZtoXConsistencyByTarget"})
    public ResponseResult<?> syncZtoXConsistencyByTarget(BigInteger partOid, String partNo, String targetXiangOids) {
        List<String> targetList = null;
        if (StringUtils.hasText(targetXiangOids)) {
            targetList = Arrays.stream(targetXiangOids.split(","))
                    .map(String::trim)
                    .filter(StringUtils::hasText)
                    .collect(Collectors.toList());
        }
        ezBomConsistencyService.syncZtoXConsistency(partOid, partNo, targetList);
        return ResponseResult.success();
    }

    @ApiOperation(value = "历史数据用,根据入库单元直接同步入库单元下所有项结构到种", notes = "根据种的bom结构创建项的bom结构", produces = "application/json; charset=utf-8")
    @GetMapping({"/syncHistoryXtoZConsistency"})
    public void syncHistoryXtoZConsistency(BigInteger rkdyOid) {
        ezBomConsistencyService.syncHistoryXtoZConsistency(rkdyOid);
    }

    @ApiOperation(value = "导出检入同步通知列表为Excel", notes = "前端反传检入时收集的非正在工作订货信息列表,导出为xlsx", produces = "application/octet-stream; charset=utf-8")
    @PostMapping({"/exportSyncNotifyList"})
    public void exportSyncNotifyList(@RequestBody List<EzCheckInSyncResultVO> dataList, HttpServletResponse response) throws IOException {
        String fileName = URLEncoder.encode("检入同步通知列表", "UTF-8");
        response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
        response.setCharacterEncoding("utf-8");
        response.setHeader("Content-disposition", "attachment;filename=" + fileName + ".xlsx");
        // 表头样式
        WriteCellStyle headWriteCellStyle = new WriteCellStyle();
        headWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        WriteFont headWriteFont = new WriteFont();
        headWriteFont.setFontHeightInPoints((short) 11);
        headWriteFont.setBold(true);
        headWriteCellStyle.setWriteFont(headWriteFont);
        // 内容样式
        WriteCellStyle contentWriteCellStyle = new WriteCellStyle();
        contentWriteCellStyle.setHorizontalAlignment(HorizontalAlignment.CENTER);
        contentWriteCellStyle.setWrapped(true);
        HorizontalCellStyleStrategy styleStrategy = new HorizontalCellStyleStrategy(headWriteCellStyle, contentWriteCellStyle);
        EasyExcel.write(response.getOutputStream(), EzCheckInSyncResultVO.class)
                .registerWriteHandler(styleStrategy)
                .sheet("通知列表")
                .doWrite(dataList == null ? Collections.emptyList() : dataList);
    }

}
