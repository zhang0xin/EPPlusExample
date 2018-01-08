declare
    type t_list is table of varchar(100);

    v_table_names t_list;
    v_field_names t_list;
    v_field_types t_list;
    v_filter nvarchar2(200) := '%';
    v_keyword nvarchar2(200) := 'lNpmh1PeRVyFYTEy7L0kSPd2bqg='; 
    v_count number(10, 0);
begin
    execute immediate 'select table_name from user_tables where table_name like :1' 
        bulk collect into v_table_names
        using v_filter;
    for i in v_table_names.first..v_table_names.last loop
        --dbms_output.put_line('Table:' || v_table_names(i));
        execute immediate 'select column_name, data_type from user_tab_columns where (table_name=:1 and data_type like :2)'
            bulk collect into v_field_names, v_field_types
            using v_table_names(i), '%CHAR%';
        if v_field_names.count > 0 then
            for j in v_field_names.first..v_field_names.last loop
                execute immediate 'select count(*) from ' || v_table_names(i) || ' where ' || v_field_names(j) || ' like :1'
                    into v_count using v_keyword;
                if (v_count > 0) then
                    dbms_output.put_line(v_table_names(i) || '  ' || v_field_names(j) || ' like ''' || v_keyword || '''  ' || v_count);
                end if;
            end loop;
        end if;
    end loop;
end;