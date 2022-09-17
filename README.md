# SampleExcelDownloadProject

fastexcel과 Apache POI 라이브러리 성능 비교를 위해 개발한 샘플 프로젝트

<h6>fastexcel 장점</h6>

* HTTP 요청에 대한 응답을 바로 내려주므로 gateway timeout error 발생하지 않습니다.
* response stream에 XLSX를 바로 작성하기 때문에 다운로드 창이 바로 뜹니다. (높은 사용성) 

<h6>fastexcel 단점</h6>

* Apache POI보다 전체적인 성능에 열세를 보입니다.
* 버전이 아직 1점대가 아니고 업데이트 속도가 느립니다.
 
<h6>Apache POI 장점</h6>

* 버전업이 꾸준히 됩니다. (현재 5점대)
* JVM Heap Memory 관리를 잘하며 소요시간 및 CPU 사용량 측면에서도 성능이 우세합니다.

<h6>Apache POI 단점</h6>

* Workbook을 생성 완료한 뒤 response stream에 내려줄 수 있으므로 대용량 엑셀의 경우 default gateway timeout 시간을 늘려줘야 합니다.
* 다운로드 창이 늦게 뜨므로 사용성 측면에서 fastexcel에 비해 열세입니다.

자세한 내용: https://jaimemin.tistory.com/2191
