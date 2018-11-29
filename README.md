### docker
- 镜像列表  
> docker images
- 删除镜像
> docker rmi `镜像ID`
- 正在运行的容器
> docker ps
- 所有容器
> docker ps -a
- 删除容器
> docker rm `容器ID`
- 导出本地镜像包
> docker save -o `文件路径+后缀`  `镜像名:版本号`
> docker save `镜像名:版本号`  >  `文件路径+后缀`
- 导入镜像包
> docker load -i `位置`
> docker load < `位置`
- 关闭容器(发送SIGTERM信号,做一些'退出前工作',再发送SIGKILL信号)
> docker stop anyesu-container
- 强制关闭容器(默认发送SIGKILL信号, 加-s参数可以发送其他信号)
> docker kill `容器ID`
- 启动容器
> docker start `容器ID`
- 重启容器
> docker restart `容器ID`
- 显示容器硬件资源使用情况
> docker stats
