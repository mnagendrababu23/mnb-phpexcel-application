<?php

declare(strict_types=1);

namespace Mnb\PHPExcel\Application\Http;

use Mnb\PHPExcel\Application\SpreadsheetApi;
use PDO;
use Throwable;

/**
 * Complete framework-neutral HTTP endpoint: routing, JSON/form parsing, CORS,
 * bearer/HMAC/CSRF authentication, rate limiting and consistent responses.
 */
final class SpreadsheetHttpEndpoint
{
    /** @param array<string,mixed> $options */
    public function __construct(private readonly SpreadsheetApi $api=new SpreadsheetApi(), private readonly array $options=[])
    {
    }

    /** @param array<string,mixed> $server @param array<string,mixed> $query @param array<string,mixed> $form @param array<string,mixed> $files */
    public function handle(array $server, array $query=[], array $form=[], array $files=[], string $rawBody='', PDO|array|string|null $pdo=null): HttpResponse
    {
        $origin=(string)($server['HTTP_ORIGIN']??'');$cors=$this->corsHeaders($origin);
        $method=strtoupper((string)($server['REQUEST_METHOD']??'GET'));
        if ($method==='OPTIONS') { return new HttpResponse(204,'',$cors+['Access-Control-Allow-Methods'=>'GET, POST, OPTIONS','Access-Control-Allow-Headers'=>'Authorization, Content-Type, X-MNB-Timestamp, X-MNB-Signature, X-CSRF-Token']); }
        try {
            $this->authenticate($server,$method,$rawBody);
            $rate=$this->rateLimit($server);
            if (!$rate['allowed']) { return HttpResponse::json(['ok'=>false,'status'=>'rate_limited','error'=>'Too many requests.'],429,$cors+$this->rateHeaders($rate)); }
            $request=$form;
            $contentType=strtolower((string)($server['CONTENT_TYPE']??''));
            if (str_contains($contentType,'application/json') && trim($rawBody)!=='') {
                $decoded=json_decode($rawBody,true,512,JSON_THROW_ON_ERROR); if(is_array($decoded)){$request=$decoded+$request;}
            }
            if ($files!==[]) { $request['file']=$files['file']??reset($files); }
            $action=$this->resolveAction($server,$query,$request);
            $allowedMethods=['status'=>['GET','POST'],'preview'=>['GET','POST'],'upload'=>['POST'],'import'=>['POST'],'import-many'=>['POST'],'export'=>['POST']];
            if (!in_array($method,$allowedMethods[$action]??['POST'],true)) { return HttpResponse::json(['ok'=>false,'status'=>'method_not_allowed','error'=>'HTTP method not allowed.'],405,$cors+['Allow'=>implode(', ',$allowedMethods[$action]??['POST'])]); }
            $result=$this->api->handle($action,$request,$pdo);$status=(int)($result['http_status']??(($result['ok']??false)?200:422));
            return HttpResponse::json($result,$status,$cors+$this->rateHeaders($rate));
        } catch (Throwable $e) {
            return HttpResponse::json(['ok'=>false,'status'=>'error','error'=>$e->getMessage()],$e instanceof \JsonException?400:401,$cors);
        }
    }

    /** Handle PHP superglobals directly. */
    public function handleGlobals(PDO|array|string|null $pdo=null): HttpResponse
    {
        return $this->handle($_SERVER,$_GET,$_POST,$_FILES,(string)file_get_contents('php://input'),$pdo);
    }

    /** @param array<string,mixed> $server */
    private function authenticate(array $server,string $method,string $body):void
    {
        $tokens=array_values((array)($this->options['bearer_tokens']??[]));$secret=(string)($this->options['hmac_secret']??'');$csrf=(string)($this->options['csrf_token']??'');
        if($tokens!==[]){$auth=(string)($server['HTTP_AUTHORIZATION']??'');$provided=str_starts_with($auth,'Bearer ')?substr($auth,7):'';$ok=false;foreach($tokens as $token){if(hash_equals((string)$token,$provided)){$ok=true;break;}}if(!$ok)throw new \RuntimeException('Invalid bearer token.');}
        if($secret!==''){$timestamp=(string)($server['HTTP_X_MNB_TIMESTAMP']??'');$signature=(string)($server['HTTP_X_MNB_SIGNATURE']??'');if(!ctype_digit($timestamp)||abs(time()-(int)$timestamp)>max(30,(int)($this->options['hmac_max_skew_seconds']??300)))throw new \RuntimeException('Expired HMAC timestamp.');$path=(string)($server['REQUEST_URI']??'/');$expected=hash_hmac('sha256',$method."\n".$path."\n".$timestamp."\n".hash('sha256',$body),$secret);if(!hash_equals($expected,$signature))throw new \RuntimeException('Invalid HMAC signature.');}
        if($csrf!==''&&!in_array($method,['GET','HEAD','OPTIONS'],true)){if(!hash_equals($csrf,(string)($server['HTTP_X_CSRF_TOKEN']??'')))throw new \RuntimeException('Invalid CSRF token.');}
    }

    /** @param array<string,mixed> $server @return array{allowed:bool,limit:int,remaining:int,retry_after:int,reset:int} */
    private function rateLimit(array $server):array
    {
        $limiter=$this->options['rate_limiter']??null;if(!$limiter instanceof RateLimiterInterface)return ['allowed'=>true,'limit'=>0,'remaining'=>0,'retry_after'=>0,'reset'=>0];
        $identity=(string)($server['HTTP_X_FORWARDED_FOR']??$server['REMOTE_ADDR']??'unknown');return $limiter->consume($identity,max(1,(int)($this->options['rate_limit']??60)),max(1,(int)($this->options['rate_window_seconds']??60)));
    }
    /** @param array{allowed:bool,limit:int,remaining:int,retry_after:int,reset:int} $rate @return array<string,string> */
    private function rateHeaders(array $rate):array{return $rate['limit']<=0?[]:['X-RateLimit-Limit'=>(string)$rate['limit'],'X-RateLimit-Remaining'=>(string)$rate['remaining'],'X-RateLimit-Reset'=>(string)$rate['reset']]+($rate['retry_after']>0?['Retry-After'=>(string)$rate['retry_after']]:[]);}
    /** @return array<string,string> */
    private function corsHeaders(string $origin):array{$allowed=(array)($this->options['allowed_origins']??[]);if($origin===''||$allowed===[])return[];if(in_array('*',$allowed,true)||in_array($origin,$allowed,true))return['Access-Control-Allow-Origin'=>in_array('*',$allowed,true)?'*':$origin,'Vary'=>'Origin'];return[];}
    /** @param array<string,mixed> $server @param array<string,mixed> $query @param array<string,mixed> $request */
    private function resolveAction(array $server,array $query,array $request):string
    {
        $action=(string)($query['action']??$request['action']??'');if($action!=='')return strtolower($action);$path=parse_url((string)($server['REQUEST_URI']??''),PHP_URL_PATH);$base=rtrim((string)($this->options['base_path']??'/spreadsheet'),'/');if(is_string($path)&&str_starts_with($path,$base.'/'))return strtolower(trim(substr($path,strlen($base)),'/'));return 'status';
    }
}
