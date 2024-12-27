@echo off

set Options=cpd --minimum-tokens=75 --ignore-annotations --ignore-identifiers --ignore-literal-sequences --ignore-literals --ignore-sequences --no-fail-on-error --no-skip-blocks -e GBK -l cpp

set BasePath=F:\newengine\dtsre
set zzorgPath=F:\zzchange\zzcode

echo running CommonLibaray...
call pmd.bat %Options% -d %BasePath%\CommonLibrary,%zzorgPath% > CommonLibaray.txt

echo running RenderEngine...
call pmd.bat %Options% -d %BasePath%\RenderEngine,%zzorgPath% > RenderEngine.txt

echo running RenderSystem...
call pmd.bat %Options% -d %BasePath%\RenderSystem,%zzorgPath% > RenderSystem.txt

echo running RenderControl...
call pmd.bat %Options% -d %BasePath%\RenderControl,%zzorgPath% > RenderControl.txt