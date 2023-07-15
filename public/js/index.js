const baseUrl ="http://localhost:3000";
  let fileInputCheck = document.getElementById('file');
  let fileInputHours = document.getElementById('file1');
  let fileListDiv = document.getElementById('fileList');
 // 获取loading元素
 let loadingElement = document.getElementById('loading');
// 获取提交元素
let submitCheckBtn = document.getElementById('submitcheck');
let submitHoursBtn = document.getElementById('submithours');
  submitCheckBtn.addEventListener("click",function(){
      uploadFiles(fileInputCheck,'mul','check');
    })

  submitHoursBtn.addEventListener("click",function(){
    uploadFiles(fileInputHours,'single','hours');
  })

  fileInputCheck.addEventListener('change', function() {
    // 清空文件列表
    fileListDiv.innerHTML = '';

    // 获取选择的文件列表
    let fileList = fileInputCheck.files;

    // 遍历文件列表并展示文件名
    for (let i = 0; i < fileList.length; i++) {
        let file = fileList[i];
        let fileName = file.name;
        // 创建一个新的<div>元素来展示文件名
        let fileDiv = document.createElement('div');
        fileDiv.textContent = fileName;
        // 将<div>元素添加到文件列表<div>中
        fileListDiv.appendChild(fileDiv);
    }
});


  function uploadFiles(dom,num,url) {
    let files = dom.files;
    let formData = new FormData();
   if(num === 'mul'){
    if(files.length == 0){
      alert('请选择文件');
      return;
     }
    if(files.length == 1){
      alert('请选择至少2个文件');
      return;
     }
   }
   
   showLoading();// 上传动画
    for (let i = 0; i < files.length; i++) {
      formData.append('files', files[i]);
    }
    var xhr = new XMLHttpRequest();
    xhr.open('POST', baseUrl+'/api/upload/'+url);
    xhr.send(formData);

    xhr.onreadystatechange = function() {
      if (xhr.readyState === XMLHttpRequest.DONE) {
        hideLoading();
        if (xhr.status == 200) {
          var uploadSuccessModal = document.getElementById('uploadSuccess');
          uploadSuccessModal.style.display = 'block';
          const timerId = setTimeout(() => {
            location.reload();
          }, 1000);
          
          
          
        } else {
          var uploadFailModal = document.getElementById('uploadFail');
          uploadFailModal.style.display = 'block';
          const timerId1 = setTimeout(() => {
            // location.reload()
          }, 1000);
        }
      }
    };
  }

 // 显示加载动画
 function showLoading() {
  loadingElement.style.display = 'flex';
}

// 隐藏加载动画
function hideLoading() {
  loadingElement.style.display = 'none';
}

// tips弹窗
function closeModal() {
    var uploadSuccessModal = document.getElementById('uploadSuccess');
    var uploadFailModal = document.getElementById('uploadFail');
    uploadSuccessModal.style.display = 'none';
    uploadFailModal.style.display = 'none';
}

// 左侧菜单的点击事件
const menuItems = document.querySelectorAll('.sidebar a');
const contentDivs = document.querySelectorAll('.content div');
menuItems.forEach(item => {
  item.addEventListener('click', () => {
    const target = item.getAttribute('data-target');
    menuItems.forEach(item => item.classList.remove('active'));
    item.classList.add('active');
    contentDivs.forEach(div => div.classList.remove('active'));
    document.getElementById(target).classList.add('active');
  });
});






