<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application;

/** Generates a dependency-free upload form and JavaScript client for AjaxUploadHandler/SpreadsheetHttpEndpoint. */
final class AjaxUploader
{
    /** @param array<string,mixed> $options */
    public static function html(string $endpoint, array $options=[]): string
    {
        $id=preg_replace('/[^A-Za-z0-9_-]/','',(string)($options['id']??'mnb-excel-upload'))?:'mnb-excel-upload';$accept=htmlspecialchars((string)($options['accept']??'.xlsx,.xls,.ods,.csv,.tsv'),ENT_QUOTES,'UTF-8');$label=htmlspecialchars((string)($options['label']??'Choose spreadsheet'),ENT_QUOTES,'UTF-8');
        return '<form id="'.$id.'" enctype="multipart/form-data"><label>'.$label.' <input name="file" type="file" accept="'.$accept.'" required></label><button type="submit">Upload</button><progress value="0" max="100" hidden></progress><pre aria-live="polite"></pre></form><script>'.self::script($id,$endpoint,$options).'</script>';
    }
    /** @param array<string,mixed> $options */
    public static function script(string $formId,string $endpoint,array $options=[]):string
    {
        $config=['formId'=>$formId,'endpoint'=>$endpoint,'headers'=>(array)($options['headers']??[]),'field'=>(string)($options['field']??'file'),'action'=>(string)($options['action']??'upload')];
        $json=json_encode($config,JSON_UNESCAPED_SLASHES|JSON_UNESCAPED_UNICODE|JSON_THROW_ON_ERROR);
        return '(function(c){const f=document.getElementById(c.formId);if(!f)return;const p=f.querySelector("progress"),o=f.querySelector("pre");f.addEventListener("submit",function(e){e.preventDefault();const file=f.querySelector("input[type=file]").files[0];if(!file)return;const d=new FormData();d.append(c.field,file);d.append("action",c.action);const x=new XMLHttpRequest();x.open("POST",c.endpoint,true);Object.entries(c.headers).forEach(([k,v])=>x.setRequestHeader(k,v));x.upload.onprogress=e=>{if(e.lengthComputable){p.hidden=false;p.value=Math.round(e.loaded/e.total*100);}};x.onload=()=>{p.hidden=true;let r;try{r=JSON.parse(x.responseText);}catch(_){r={ok:false,error:x.responseText};}o.textContent=JSON.stringify(r,null,2);f.dispatchEvent(new CustomEvent("mnb-upload-complete",{detail:r}));};x.onerror=()=>{o.textContent="Upload failed";};x.send(d);});})('.$json.');';
    }
}
