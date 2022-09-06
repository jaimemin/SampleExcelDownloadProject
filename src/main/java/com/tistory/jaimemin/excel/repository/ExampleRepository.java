package com.tistory.jaimemin.excel.repository;

import com.tistory.jaimemin.excel.vo.ExampleVO;
import org.springframework.data.domain.Page;
import org.springframework.data.domain.Pageable;
import org.springframework.data.jpa.repository.JpaRepository;
import org.springframework.stereotype.Repository;
import org.springframework.transaction.annotation.Transactional;

@Repository
@Transactional(readOnly = true)
public interface ExampleRepository extends JpaRepository<ExampleVO, Long> {

    Page<ExampleVO> findAll(Pageable pageable);
}
