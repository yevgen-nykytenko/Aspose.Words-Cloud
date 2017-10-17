del /S /Q "c:\tmp\csharp"
java -jar SDKs\codegen\swagger-codegen-cli-2.3.0.jar generate -i SDKs\spec\asposeforcloud_word_csharp.json -l csharp -t SDKs\codegen\Templates\csharp -o c:/tmp/csharp/ -c SDKs\codegen\config.json 

SDKs\codegen\Tools\SplitCSharpCodeFile.exe C:\tmp\csharp\src\Aspose.Words.Cloud.Sdk\Api\WordsApi.cs C:\tmp\csharp\src\Aspose.Words.Cloud.Sdk\Model\Requests\

del /S /Q "SDKs\Aspose.Words-Cloud-SDK-for-.NET\Aspose.Words.Cloud.Sdk\Model"
del /S /Q "SDKs\Aspose.Words-Cloud-SDK-for-.NET\Aspose.Words.Cloud.Sdk\Api"

xcopy "c:\tmp\csharp\src\Aspose.Words.Cloud.Sdk\Model" "SDKs\Aspose.Words-Cloud-SDK-for-.NET\Aspose.Words.Cloud.Sdk\Model" /E
xcopy "c:\tmp\csharp\src\Aspose.Words.Cloud.Sdk\Api" "SDKs\Aspose.Words-Cloud-SDK-for-.NET\Aspose.Words.Cloud.Sdk\Api" /E










