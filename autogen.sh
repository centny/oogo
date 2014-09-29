if [ `uname` = "Darwin" ];then
 glibtoolize -f -c
else
 libtoolize -f -c
fi
autoheader
touch NEWS README AUTHORS ChangeLog
automake --add-missing
automake
echo 1
aclocal -I m4
autoconf
