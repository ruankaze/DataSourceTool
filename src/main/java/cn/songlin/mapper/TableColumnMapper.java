package cn.songlin.mapper;

import cn.songlin.entity.TableColumn;
import org.apache.ibatis.annotations.Param;
import org.apache.ibatis.annotations.Select;
import tk.mybatis.mapper.common.Mapper;

import java.util.List;

public interface TableColumnMapper extends Mapper<TableColumn> {


    /**
     * 查询指定数据库和表的表结构
     *
     * @param dataSourceName
     * @param tableName
     * @return
     * @author liusonglin
     * @date 2018年7月25日
     */

    List<TableColumn> getTableColumn(@Param("dataSourceName") String dataSourceName, @Param("tableName") String tableName);


    /**
     * 查询数据所有表的表结构
     *
     * @param dataSourceName
     * @return
     * @author liusonglin
     * @date 2018年7月25日
     */

    List<TableColumn> getAllTableColumn(@Param("dataSourceName") String dataSourceName);

    String getTabComment(@Param("dataSourceName") String dataSourceName, @Param("tableName") String tableName);

}