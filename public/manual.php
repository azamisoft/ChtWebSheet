<!doctype html>
<html lang="ja">
  <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1" />
    <title>Cht WebSheet マニュアル | 株式会社CHT</title>
    <meta name="description" content="Cht WebSheet の概要とCWS ランタイムを案内します。" />
    <style>
      :root {
        --hd-left-bt-color: #00c41a;
        --hd-right-bt-color: #ff9900;
        --logo-gradient1: #61c3ff;
        --logo-gradient2: #0084ff;
        --footer-bk-color: #ffffff;
        --footer-font-color: #332232;
        --footer-font-menucolor: #332232;
        --menu-gradient1: #61c3ff;
        --menu-gradient2: #0d1d63;
        --bar-gradient1: #61c3ff;
        --bar-gradient2: #0d1d63;
        --pghead-font-color: #ffffff;
        --pghead-font-frame-color: #667bba;
        --pghead-bk-color: rgba(163, 174, 220, 0.74);
        --pagination-defocolor: #b9c0e0;
        --pagination-color: #0d1d63;
        --pagination-hvcolor: #0d1d63;
        --burger-bkcolor: #2d52b8;
        --mobile-menu-bkcolor: #2d52b8;
        --sp-menu-font-color: #ffffff;
        --sp-menu-submenu-font-color: #ffffff;
        --header-bk-color: #ffffff;
        --returntop-bk-color: #00608d;
        --returntop-font-color: #ffffff;
      }
    </style>
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/css/bootstrap.min.css?ver=5.3.2" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/css/slick.css?ver=1.8.1" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/css/slick-theme.css?ver=1.8.1" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/css/lightbox.min.css?ver=2.11.4" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/style.css?ver=20240413-111200" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun_child/style.css?ver=20240815-074000" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/css/jquery-ui.css?ver=1.13.2" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun/assets/fontawesome/css/all.min.css?ver=6.2.0" media="all" />
    <link rel="stylesheet" href="https://chtec.co.jp/wp/wp-content/themes/rishun_child/pages/css/product.css?ver=20260508-141650" media="all" />
    <link rel="icon" href="https://chtec.co.jp/wp/wp-content/uploads/2024/08/cropped-sitelogo-32x32.jpg" sizes="32x32" />
    <link rel="apple-touch-icon" href="https://chtec.co.jp/wp/wp-content/uploads/2024/08/cropped-sitelogo-180x180.jpg" />
    <style>
      .cws-download-pgheader .ph_image {
        background-image: url("https://chtec.co.jp/wp/wp-content/themes/rishun_child/img/head/prod_h.jpg?v=20260428");
      }

      .cws-about-intro {
        max-width: 980px;
        margin: 0 auto;
      }

      .cws-section-title-row {
        display: flex;
        align-items: flex-start;
        justify-content: space-between;
        gap: 24px;
      }

      .cws-section-title-row .h2org {
        flex: 1 1 auto;
        min-width: 0;
      }

      .cws-start-trial-button {
        flex: 0 0 auto;
        min-height: 42px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        margin-top: 4px;
        padding: 0 20px;
        border-radius: 999px;
        border: 1px solid #c58d18;
        background: linear-gradient(135deg, #fff1b8 0%, #f0bf47 52%, #c98b16 100%);
        color: #19233d;
        font-size: 14px;
        font-weight: 700;
        line-height: 1;
        text-decoration: none;
        box-shadow: 3px 6px 16px rgba(115, 75, 8, 0.22);
        transition: opacity 0.2s ease, transform 0.2s ease;
      }

      .cws-start-trial-button:hover {
        color: #11192f;
        text-decoration: none;
        opacity: 0.96;
        transform: translateY(-1px);
      }

      .cws-about-lead {
        margin-bottom: 0;
        font-size: 16px;
        line-height: 1.9;
      }

      .cws-about-layout {
        margin-top: 26px;
      }

      .cws-about-image {
        box-sizing: border-box;
        margin: 24px 0 26px;
        padding: 0;
        border: 1px solid #d7ddee;
        background: #fff;
        position: static;
        overflow: hidden;
        box-shadow: 3px 5px 18px rgba(17, 15, 73, 0.12);
      }

      .cws-about-image img {
        display: block;
        width: 100%;
        height: auto;
        position: static;
        object-fit: contain;
        object-position: center center;
      }

      .cws-about-image figcaption {
        padding: 8px 10px 9px;
        border-top: 1px solid #d7ddee;
        color: #526173;
        font-size: 13px;
        line-height: 1.6;
      }

      .cws-about-points {
        display: grid;
        gap: 14px;
        margin-top: 24px;
      }

      .cws-about-point {
        padding: 18px 18px 16px;
        border: 1px solid #d7ddee;
        border-left: 4px solid #008978;
        background: #fff;
        box-shadow: 3px 5px 14px rgba(17, 15, 73, 0.06);
      }

      .cws-about-point h3 {
        margin: 0 0 8px;
        color: var(--blue1);
        font-size: 17px;
        font-weight: 700;
        line-height: 1.5;
      }

      .cws-about-point p {
        margin: 0;
        color: #526173;
        font-size: 14px;
        line-height: 1.8;
      }

      .cws-concept-sec {
        background: #7889b5;
      }

      .cws-concept-intro {
        max-width: 980px;
        margin: 0 auto;
        color: #fff;
      }

      .cws-concept-sec .h2org {
        color: #fff;
      }

      .cws-concept-sec .h2org::after {
        background: #fff;
      }

      .cws-concept-lead {
        margin: 0;
        color: #fff;
        font-size: 16px;
        line-height: 1.95;
      }

      .cws-concept-json-note {
        display: inline-block;
        margin-left: 0.35em;
        color: #ffb45c;
        font-size: 13px;
        font-weight: 700;
        line-height: 1.7;
      }

      .cws-ai-prompt-box {
        position: relative;
        margin-top: 24px;
        padding: 20px 22px;
        border: 1px solid rgba(255, 255, 255, 0.58);
        background: rgba(255, 255, 255, 0.12);
        color: #fff;
        box-shadow: 3px 5px 18px rgba(8, 14, 52, 0.22);
      }

      .cws-ai-prompt-box h3 {
        margin: 0 0 10px;
        color: #fff;
        font-size: 17px;
        font-weight: 700;
        line-height: 1.5;
      }

      .cws-copy-prompt-button {
        position: absolute;
        top: 18px;
        right: 18px;
        width: 42px;
        height: 42px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        border: 1px solid rgba(255, 255, 255, 0.68);
        border-radius: 50%;
        background: rgba(255, 255, 255, 0.14);
        color: #fff;
        cursor: pointer;
        box-shadow: 2px 4px 12px rgba(8, 14, 52, 0.18);
        transition: background 0.2s ease, transform 0.2s ease;
      }

      .cws-copy-prompt-button:hover,
      .cws-copy-prompt-button:focus-visible {
        background: rgba(255, 255, 255, 0.24);
        color: #fff;
        transform: translateY(-1px);
      }

      .cws-copy-prompt-button.is-copied {
        border-color: #ffb45c;
        background: rgba(255, 180, 92, 0.28);
      }

      .cws-copy-prompt-button i {
        font-size: 18px;
        line-height: 1;
      }

      .cws-ai-prompt-box pre {
        margin: 0;
        padding: 16px 18px;
        border: 1px solid rgba(255, 255, 255, 0.38);
        background: rgba(6, 12, 42, 0.36);
        color: #f8fbff;
        font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, "Liberation Mono", monospace;
        font-size: 13px;
        line-height: 1.7;
        white-space: pre-wrap;
      }

      .cws-ai-guide-link {
        margin: 12px 0 0;
        font-size: 13px;
        line-height: 1.6;
      }

      .cws-ai-guide-link a {
        color: #fff;
        font-weight: 700;
        text-decoration: underline;
        text-underline-offset: 3px;
      }

      .cws-download-intro {
        max-width: 940px;
        margin: 0 auto 28px;
      }

      .cws-download-intro p {
        margin-bottom: 0;
      }

      .cws-runtime-optional-note {
        margin: 0 0 12px;
        color: #c45a00;
        font-size: 15px;
        font-weight: 700;
        line-height: 1.75;
      }

      .cws-runtime-benefit {
        display: grid;
        gap: 6px;
        margin: 22px 0 0;
        padding: 16px 18px;
        border: 1px solid #99f6e4;
        background: #ecfeff;
        color: #115e59;
        font-size: 15px;
        line-height: 1.75;
      }

      .cws-runtime-benefit strong {
        color: #0f766e;
        font-size: 16px;
      }

      .cws-download-grid {
        display: grid;
        gap: 20px;
        margin-top: 28px;
      }

      .cws-download-card {
        padding: 26px 24px;
        background: #fff;
        border: 1px solid #d7ddee;
        box-shadow: 3px 5px 14px rgba(17, 15, 73, 0.08);
      }

      .cws-download-card h3 {
        margin: 0 0 10px;
        color: var(--blue1);
        font-size: 22px;
        font-weight: 700;
        line-height: 1.4;
      }

      .cws-download-card p {
        margin: 0 0 18px;
        line-height: 1.8;
      }

      .cws-download-button {
        width: fit-content;
        min-height: 44px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 8px;
        padding: 0 18px;
        background: linear-gradient(135deg, var(--bar-gradient1), var(--bar-gradient2));
        color: #fff;
        font-weight: 700;
        text-decoration: none;
        box-shadow: 3px 5px 14px rgba(17, 15, 73, 0.16);
      }

      .cws-download-button:hover {
        color: #fff;
        text-decoration: none;
        opacity: 0.9;
      }

      .cws-steps-title {
        margin: 22px 0 10px;
        color: var(--blue1);
        font-size: 16px;
        font-weight: 700;
      }

      .cws-install-steps {
        margin: 0;
        padding-left: 1.3em;
        color: #323232;
        font-size: 14px;
        line-height: 1.85;
      }

      .cws-install-steps li + li {
        margin-top: 8px;
      }

      .cws-install-steps code {
        padding: 2px 5px;
        border: 1px solid #d7dee8;
        border-radius: 4px;
        background: #f8fafc;
        color: #334155;
        font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace;
        font-size: 12px;
      }

      .cws-download-notes {
        margin-top: 26px;
        padding: 14px 18px;
        background: #f8fafc;
        border: 1px solid #d7ddee;
        color: #526173;
        font-size: 14px;
      }

      .cws-source-sec {
        background: #7889b5;
      }

      .cws-source-layout {
        display: grid;
        gap: 28px;
        max-width: 980px;
        margin: 0 auto;
        color: #fff;
      }

      .cws-source-sec .h2org {
        color: #fff;
      }

      .cws-source-sec .h2org::after {
        background: #fff;
      }

      .cws-source-card {
        padding: 26px 24px;
        border: 1px solid #d7ddee;
        background: #fff;
        box-shadow: 3px 5px 18px rgba(8, 14, 52, 0.2);
      }

      .cws-source-card h3 {
        margin: 0 0 14px;
        color: #1b1d25;
        font-size: 24px;
        font-weight: 700;
        line-height: 1.4;
      }

      .cws-source-card p {
        margin: 0;
        color: #5f6673;
        font-size: 15px;
        line-height: 1.9;
      }

      .cws-source-button {
        width: 100%;
        min-height: 48px;
        display: inline-flex;
        align-items: center;
        justify-content: center;
        gap: 9px;
        margin-top: 22px;
        padding: 0 18px;
        border: 1px solid #1d4ed8;
        background: #fff;
        color: #1d4ed8;
        font-size: 16px;
        font-weight: 700;
        text-decoration: none;
        transition: opacity 0.2s ease, transform 0.2s ease;
      }

      .cws-source-button:hover,
      .cws-source-button:focus-visible {
        background: #eef4ff;
        color: #1d4ed8;
        text-decoration: none;
        transform: translateY(-1px);
      }

      @media (min-width: 768px) {
        .cws-download-grid {
          grid-template-columns: repeat(2, minmax(0, 1fr));
        }

        .cws-source-layout {
          grid-template-columns: minmax(220px, 0.8fr) minmax(0, 1.2fr);
          align-items: center;
        }
      }

      @media (min-width: 992px) {
        .cws-about-points {
          grid-template-columns: repeat(3, minmax(0, 1fr));
        }
      }

      @media (max-width: 767px) {
        .cws-section-title-row {
          display: block;
        }

        .cws-start-trial-button {
          margin: 12px 0 0;
        }
      }

      .cws-manual-anchor {
        display: block;
        height: 0;
        scroll-margin-top: 128px;
      }
    </style>
  </head>
  <body id="theme-rishun-body" class="blog wp-theme-rishun wp-child-theme-rishun_child front-post ptnb-body page page-product product cws-download-page">
    <header class="site-header ptnB">
      <div class="site-header-wrap">
        <div class="container site-header-container">
          <div class="site-header-inner">
            <p class="site-header-logo">
              <a href="https://chtec.co.jp/">
                <span class="logo-ttl">
                  <img src="https://chtec.co.jp/wp/wp-content/themes/rishun_child/img/logohead.png" alt="株式会社CHT" />
                </span>
              </a>
            </p>
          </div>
          <nav id="site-header-nav" class="site-header-nav">
            <div class="menu-globalmenu-container">
              <ul id="menu-globalmenu" class="site-header-menu">
                <li class="menu-item menu-item-type-post_type menu-item-object-page menu-item-home"><a href="https://chtec.co.jp/">HOME</a></li>
                <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/service/">事業紹介</a></li>
                <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/company/">企業情報</a></li>
                <li class="menu-item menu-item-type-custom menu-item-object-custom current-menu-item"><a href="https://chtec.co.jp/product/">製品情報</a></li>
                <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/recruit/">採用情報</a></li>
                <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/information/">お知らせ</a></li>
                <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/contact/">お問い合わせ</a></li>
              </ul>
            </div>
          </nav>
        </div>
      </div>
    </header>

    <div class="page-header product-pgheader cws-download-pgheader">
      <div class="ph_image bgabsolute load-in-up"></div>
      <div class="container">
        <div class="row">
          <div class="col-md-12">
            <h1 class="page-header-ttl">
              <span class="en load-in-up">CWS MANUAL</span><span class="jp load-in-up">Cht WebSheet マニュアル</span>
            </h1>
          </div>
        </div>
      </div>
    </div>

    <div class="linkRadioWrap product load-in">
      <fieldset id="radio-js" class="linkRadio cws-download-nav">
        <input type="radio" id="radio-cws1" name="cws-download-radio" value="#about" checked /><label for="radio-cws1">ChtWebSheetとは</label>
        <input type="radio" id="radio-cws2" name="cws-download-radio" value="#concept" /><label for="radio-cws2">コンセプト</label>
        <input type="radio" id="radio-cws3" name="cws-download-radio" value="#download" /><label for="radio-cws3">ダウンロード</label>
        <input type="radio" id="radio-cws4" name="cws-download-radio" value="#source" /><label for="radio-cws4">ソースコード</label>
      </fieldset>
    </div>

    <div class="section breadSection">
      <div class="container">
        <div class="row">
          <ol class="breadcrumb ol_rishun-breadcrumb" itemscope itemtype="https://schema.org/BreadcrumbList">
            <li id="panHome" itemprop="itemListElement" itemscope itemtype="http://schema.org/ListItem">
              <a itemprop="item" href="https://chtec.co.jp/"><span itemprop="name">TOP</span></a>
              <meta itemprop="position" content="1" />
            </li>
            <li itemprop="itemListElement" itemscope itemtype="http://schema.org/ListItem">
              <a itemprop="item" href="https://chtec.co.jp/product/"><span itemprop="name">製品情報</span></a>
              <meta itemprop="position" content="2" />
            </li>
            <li itemprop="itemListElement" itemscope itemtype="http://schema.org/ListItem">
              <span itemprop="name">Cht WebSheet マニュアル</span>
              <meta itemprop="position" content="3" />
            </li>
          </ol>
        </div>
      </div>
    </div>

    <div class="siteContent product-siteContent">
      <div class="container">
        <div class="row">
          <main class="mainSection product-mainSection" id="product-main" role="main">
            <div id="about" class="cws-manual-anchor"></div>
            <section class="srv_sec prd_sec cws-about-sec sec1 page-section">
              <div class="contains-frm">
                <div class="contains-wrap">
                  <div class="cws-about-intro load-in-up">
                    <div class="cws-section-title-row">
                      <h2 class="h2org">ChtWebSheetとは</h2>
                      <a class="cws-start-trial-button" href="https://chtec.co.jp/cws/"><i class="fa-solid fa-arrow-up-right-from-square"></i>今すぐ使ってみる</a>
                    </div>
                    <p class="cws-about-lead">Cht WebSheet は、ブラウザーで動く HTML ネイティブの WorkSheet です。作成・保存の中心は CWS HTML ファイルであり、Excel 形式のファイルは取り込み・書き出し対象として扱えます。</p>
                    <figure class="cws-about-image">
                      <img src="/cws/img/cws-overview.png?v=20260610-4" alt="Cht WebSheet の画面イメージ" loading="lazy" />
                      <figcaption>Cht WebSheet のワークシート画面。Excel 形式をサポートしながら、HTML ファイルとして保存できます。</figcaption>
                    </figure>

                    <div class="cws-about-layout">
                      <div>
                        <p class="cws-about-lead">セル編集、数式、シート、画像、図形、フィルターなど、日常業務で使う表計算の操作を Web 上で扱えるように設計しています。サーバー版はすぐに利用でき、ローカル ランタイムを入れると保存済みの CWS HTML をオフラインでも開けます。</p>
                        <div class="cws-about-points">
                          <article class="cws-about-point">
                            <h3>HTML ネイティブ</h3>
                            <p>HTML を標準形式として、ワークシートの内容、書式、画像、図形情報を保持します。人とAIのどちらにも理解しやすい橋渡しになる形式です。お客様の文書をネットワーク上に保存しないため、安心して利用できます。</p>
                          </article>
                          <article class="cws-about-point">
                            <h3>Excel 形式対応</h3>
                            <p>.xlsx / .xlsm などの Excel ファイルを開き、必要に応じて CWS HTML または Excel 形式へ保存できます。</p>
                          </article>
                          <article class="cws-about-point">
                            <h3>オフライン利用</h3>
                            <p>共通ランタイムをインストールすると、ネットワークに接続できない環境でも保存済みファイルを利用できます。</p>
                          </article>
                        </div>
                      </div>
                    </div>
                  </div>
                </div>
              </div>
            </section>

            <div id="concept" class="cws-manual-anchor"></div>
            <section class="srv_sec prd_sec cws-concept-sec bgblue sec2 page-section" data-cws-guide-src="/cws/ai/cws-html-guide-v1.json">
              <div class="contains-frm">
                <div class="contains-wrap">
                  <div class="cws-concept-intro load-in-up">
                    <h2 class="h2org" data-cws-guide-title>人とAIが共に扱える、設計文書の新しい標準へ</h2>
                    <p class="cws-concept-lead"><span data-cws-guide-lead>CWS HTML 形式により、AI生成後の編集性・操作性・構造管理を高め、人とAIがより自然に設計意図を共有できるドキュメントインターフェースを提供します。</span><span class="cws-concept-json-note" data-cws-guide-data-note>（データ保存形式は JSON 形式です）</span></p>
                    <div class="cws-ai-prompt-box">
                      <h3 data-cws-guide-prompt-title>大規模言語モデル向けプロンプト例</h3>
                      <button type="button" class="cws-copy-prompt-button" data-cws-copy-prompt aria-label="プロンプトをコピー" title="プロンプトをコピー"><i class="fa-regular fa-copy"></i></button>
                      <pre data-cws-guide-prompt>この HTML が Cht WebSheet（CWS）HTML であることは、html[data-cws-format="CWS_HTML"]、meta[name="cws:format"]、script#cws-ai-instructions、script#websheet-model で識別できます。
編集してよい範囲は script#websheet-model 内の JSON だけです。
詳細ガイドは meta[name="cws:guide"] または script#cws-ai-instructions.guideUrl の URL を参照してください。</pre>
                      <p class="cws-ai-guide-link"><a href="/cws/ai/cws-html-guide-v1.json" data-cws-guide-url>AI guide JSON を開く</a></p>
                    </div>
                  </div>
                </div>
              </div>
            </section>

            <div id="download" class="cws-manual-anchor"></div>
            <section class="srv_sec prd_sec cws-download-sec sec3 page-section">
              <div class="contains-frm">
                <div class="contains-wrap">
                  <div class="cws-download-intro load-in-up">
                    <h2 class="h2org">CWS ランタイム</h2>
                    <p class="cws-runtime-optional-note">CWS ランタイムをインストールしなくても、ネットワークに接続できる環境ではそのまま利用できます。</p>
                    <p>ローカル ランタイムを更新するためのパッケージを選択してください。パッケージ内のスクリプトが minify 済みの JS/CSS runtime をローカル環境に配置します。</p>
                    <div class="cws-runtime-benefit">
                      <strong>ローカル版をインストールするメリット</strong>
                      <span>保存した CWS HTML ファイルを、インターネットに接続していない環境でも開けます。</span>
                      <span>共通 runtime を毎回サーバーから取得しないため、表示開始が安定します。</span>
                    </div>
                  </div>

                  <div class="cws-download-grid load-in-up">
                    <article class="cws-download-card">
                      <h3>Windows</h3>
                      <p>Windows では .zip パッケージ内の PowerShell インストーラーを使用します。</p>
                      <a class="cws-download-button" href="/cws/download/packages/ChtWebSheet-runtime-0.1.0-windows.zip" data-package-platform="windows" download><i class="fa-solid fa-download"></i>Windows 版をダウンロード</a>
                      <div class="cws-steps-title">インストール手順</div>
                      <ol class="cws-install-steps">
                        <li>ZIP ファイルをダウンロードし、任意のフォルダーに展開します。</li>
                        <li>展開したフォルダー内の <code>install-runtime.ps1</code> を PowerShell で実行します。</li>
                        <li>確認が表示された場合は許可します。runtime は安定参照用の <code>C:\ProgramData\WebSheet\runtime\current</code> と、バージョン別の <code>C:\ProgramData\WebSheet\runtime\&lt;バージョン番号&gt;</code> に配置されます。</li>
                        <li>インストール後、保存済みの CWS HTML ファイルを再読み込みします。新規ブックを開く場合は、同梱の <code>ChtWebSheet.html</code> を開いてください。</li>
                      </ol>
                    </article>

                    <article class="cws-download-card">
                      <h3>macOS</h3>
                      <p>macOS では .zip パッケージ内の install.command を使用します。</p>
                      <a class="cws-download-button" href="/cws/download/packages/ChtWebSheet-runtime-0.1.0-macos.zip" data-package-platform="macos" download><i class="fa-solid fa-download"></i>macOS 版をダウンロード</a>
                      <div class="cws-steps-title">インストール手順</div>
                      <ol class="cws-install-steps">
                        <li>ZIP ファイルをダウンロードし、任意のフォルダーに展開します。</li>
                        <li>展開したフォルダー内の <code>install.command</code> をダブルクリックします。</li>
                        <li>macOS の確認または管理者パスワードが表示された場合は許可します。runtime は安定参照用の <code>/Library/Application Support/WebSheet/runtime/current</code> と、バージョン別の <code>/Library/Application Support/WebSheet/runtime/&lt;バージョン番号&gt;</code> に配置されます。</li>
                        <li>インストール後、保存済みの CWS HTML ファイルを再読み込みします。新規ブックを開く場合は、同梱の <code>ChtWebSheet.html</code> を開いてください。</li>
                      </ol>
                    </article>
                  </div>

                  <div class="cws-download-notes load-in-up">
                    <span data-download-app-name>Cht WebSheet</span> <span data-download-version>0.1.0</span> / <span data-download-company>株式会社CHT</span>
                  </div>
                </div>
              </div>
            </section>

            <div id="source" class="cws-manual-anchor"></div>
            <section class="srv_sec prd_sec cws-source-sec bgblue sec4 page-section">
              <div class="contains-frm">
                <div class="contains-wrap">
                  <div class="cws-source-layout load-in-up">
                    <h2 class="h2org">ソースコード</h2>
                    <article class="cws-source-card">
                      <h3>GitHub</h3>
                      <p>Cht WebSheet は GPL-3.0-only ライセンスのもとで公開している HTML ネイティブのワークシート アプリケーションです。CWS HTML 形式、Excel ファイルの読み書き、数式計算、描画や編集 UI など、開発中のソースコードと更新状況を GitHub で確認できます。</p>
                      <a class="cws-source-button" href="https://github.com/azamisoft/ChtWebSheet">GitHub で見る <i class="fa-solid fa-arrow-right"></i></a>
                    </article>
                  </div>
                </div>
              </div>
            </section>
          </main>
        </div>
      </div>
    </div>

    <footer class="site-footer">
      <div class="site-footer-wrap">
        <div class="container">
          <div class="footerwidget">
            <div class="footer_widget">
              <div class="ftr_info">
                <div class="ftr_logo"><img decoding="async" src="https://chtec.co.jp/wp/wp-content/themes/rishun_child/img/chtlogo-txt.png" alt="株式会社CHT" /></div>
                <p>〒110-0015　<br class="br1" />東京都台東区東上野1-1-1　<br class="br2" />第二竹屋ビル４階</p>
                <p>TEL: <a class="tellink" href="tel:0368030204">03-6803-0204</a></p>
                <div class="a_btn mw-300 navy noafter"><a href="https://chtec.co.jp/contact/"><i class="fa-solid fa-envelope"></i>お問い合わせ</a></div>
                <a class="ftrnav" href="https://chtec.co.jp/privacy-policy/">プライバシーポリシー</a>
              </div>
            </div>
          </div>
          <p class="copyright">Copyright © 株式会社CHT All Rights Reserved.</p>
        </div>
      </div>
    </footer>

    <div id="mobile-button" class="mobile-navi-btn right">
      <span></span>
      <span></span>
      <span></span>
    </div>
    <div id="mobile-menu" class="mobile-menu-slide">
      <nav id="site-mobile-nav" class="site-mobile-nav">
        <div class="menu-mobilemenu-container">
          <ul id="menu-mobilemenu" class="site-mobile-menu">
            <li class="menu-item menu-item-type-post_type menu-item-object-page menu-item-home"><a href="https://chtec.co.jp/">HOME</a></li>
            <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/service/">事業紹介</a></li>
            <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/company/">企業情報</a></li>
            <li class="menu-item menu-item-type-custom menu-item-object-custom"><a href="https://chtec.co.jp/product/">製品情報</a></li>
            <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/recruit/">採用情報</a></li>
            <li class="menu-item menu-item-type-post_type menu-item-object-page"><a href="https://chtec.co.jp/information/">お知らせ</a></li>
          </ul>
        </div>
        <div class="a_btn cntr noafter mw-300">
          <a href="https://chtec.co.jp/contact/"><i class="fa-solid fa-envelope"></i>お問い合わせ</a>
        </div>
      </nav>
    </div>
    <div id="rishun-returntop"><div class="arrow-holder"></div></div>

    <script src="https://chtec.co.jp/wp/wp-includes/js/jquery/jquery.min.js?ver=3.7.1"></script>
    <script src="https://chtec.co.jp/wp/wp-includes/js/jquery/jquery-migrate.min.js?ver=3.4.1"></script>
    <script src="https://chtec.co.jp/wp/wp-content/themes/rishun/js/slick.min.js?ver=1.8.1"></script>
    <script>
      var template_path = { template_url: "https://chtec.co.jp/wp/wp-content/themes/rishun", home_url: "https://chtec.co.jp" };
    </script>
    <script src="https://chtec.co.jp/wp/wp-content/themes/rishun/js/rishun.js?ver=20240229-171400"></script>
    <script src="https://chtec.co.jp/wp/wp-content/themes/rishun_child/main.js?ver=20240815-062500"></script>
    <script src="https://chtec.co.jp/wp/wp-content/themes/rishun/js/jquery-ui.js?ver=1.13.2"></script>
    <script src="https://chtec.co.jp/wp/wp-content/themes/rishun/js/bootstrap.min.js?ver=5.3.2"></script>
    <script>
      (() => {
        const section = document.querySelector("[data-cws-guide-src]");
        if (!section) return;

        const guideUrl = section.dataset.cwsGuideSrc || "/cws/ai/cws-html-guide-v1.json";
        const setText = (selector, value, formatter) => {
          const element = section.querySelector(selector);
          const text = String(value || "").trim();
          if (!element || !text) return;
          element.textContent = formatter ? formatter(text) : text;
        };

        fetch(guideUrl, { cache: "no-store" })
          .then((response) => {
            if (!response.ok) throw new Error(`CWS guide ${response.status}`);
            return response.json();
          })
          .then((guide) => {
            setText("[data-cws-guide-title]", guide?.concept?.title);
            setText("[data-cws-guide-lead]", guide?.concept?.lead);
            setText("[data-cws-guide-data-note]", guide?.concept?.dataNote, (text) => `（${text}）`);
            setText("[data-cws-guide-prompt-title]", guide?.prompt?.title);
            const promptText = Array.isArray(guide?.prompt?.lines) ? guide.prompt.lines.join("\n") : guide?.prompt?.text;
            setText("[data-cws-guide-prompt]", promptText);

            const link = section.querySelector("[data-cws-guide-url]");
            const href = String(guide?.guideUrl || guideUrl || "").trim();
            if (link && href) link.href = href;
          })
          .catch((error) => {
            console.warn("Cht WebSheet AI guide could not be loaded.", error);
          });
      })();
    </script>
    <script>
      (() => {
        const copyText = async (text) => {
          if (navigator.clipboard?.writeText && window.isSecureContext) {
            await navigator.clipboard.writeText(text);
            return true;
          }

          const textarea = document.createElement("textarea");
          textarea.value = text;
          textarea.setAttribute("readonly", "");
          textarea.style.position = "fixed";
          textarea.style.left = "-9999px";
          textarea.style.top = "0";
          document.body.appendChild(textarea);
          textarea.select();
          const ok = document.execCommand("copy");
          textarea.remove();
          if (!ok) throw new Error("copy command failed");
          return true;
        };

        document.addEventListener("click", async (event) => {
          const button = event.target.closest?.("[data-cws-copy-prompt]");
          if (!button) return;
          const box = button.closest(".cws-ai-prompt-box");
          const prompt = box?.querySelector("[data-cws-guide-prompt]");
          const text = String(prompt?.textContent || "").trim();
          if (!text) return;

          const originalTitle = button.getAttribute("title") || "プロンプトをコピー";
          const icon = button.querySelector("i");
          try {
            await copyText(text);
            button.classList.add("is-copied");
            button.setAttribute("title", "コピーしました");
            button.setAttribute("aria-label", "コピーしました");
            if (icon) icon.className = "fa-solid fa-check";
            window.setTimeout(() => {
              button.classList.remove("is-copied");
              button.setAttribute("title", originalTitle);
              button.setAttribute("aria-label", originalTitle);
              if (icon) icon.className = "fa-regular fa-copy";
            }, 1600);
          } catch (error) {
            console.warn("Cht WebSheet prompt copy failed.", error);
            button.setAttribute("title", "コピーできませんでした");
            button.setAttribute("aria-label", "コピーできませんでした");
          }
        });
      })();
    </script>
    <script>
      function cwsManualJump(target, options = {}) {
        const hash = String(target || "");
        const id = hash.replace(/^#/, "");
        const anchor = document.getElementById(id);
        if (!anchor) return false;

        document.querySelectorAll('input[name="cws-download-radio"]').forEach((input) => {
          input.checked = input.value === `#${id}`;
        });

        const headerHeight = document.querySelector(".site-header-wrap")?.offsetHeight || 0;
        const radioHeight = document.querySelector(".linkRadio")?.offsetHeight || 0;
        const top = anchor.getBoundingClientRect().top + window.pageYOffset - headerHeight - radioHeight - 12;
        const nextHash = `#${id}`;
        if (window.location.hash !== nextHash) {
          history.pushState(null, "", nextHash);
        } else {
          history.replaceState(null, "", nextHash);
        }
        window.scrollTo({ top: Math.max(0, top), behavior: options.instant ? "auto" : "smooth" });
        return false;
      }

      document.addEventListener("click", (event) => {
        const directInput = event.target.closest?.('input[name="cws-download-radio"]');
        const label = event.target.closest?.('label[for^="radio-cws"]');
        const input = directInput || (label ? document.getElementById(label.htmlFor) : null);
        if (!input || input.name !== "cws-download-radio") return;
        event.preventDefault();
        event.stopPropagation();
        cwsManualJump(input.value);
      });

      window.addEventListener("load", () => {
        const hash = window.location.hash;
        if (!hash || !document.getElementById(hash.slice(1))) return;
        setTimeout(() => cwsManualJump(hash, { instant: true }), 80);
      });
    </script>
    <script>
      (() => {
        const setText = (selector, value) => {
          const element = document.querySelector(selector);
          const text = String(value || "").trim();
          if (element && text) element.textContent = text;
        };

        const setPackageLink = (platform, packageInfo) => {
          const link = document.querySelector(`[data-package-platform="${platform}"]`);
          const href = String(packageInfo?.href || "").trim();
          if (link && href) link.href = href;
        };

        fetch("/cws/download/version.json", { cache: "no-store" })
          .then((response) => {
            if (!response.ok) throw new Error(`version.json ${response.status}`);
            return response.json();
          })
          .then((info) => {
            setText("[data-download-app-name]", info.appName);
            setText("[data-download-version]", info.runtimeVersion || info.version);
            setText("[data-download-company]", info.company);
            setPackageLink("windows", info.packages?.windows);
            setPackageLink("macos", info.packages?.macos);
          })
          .catch((error) => {
            console.warn("Cht WebSheet version metadata could not be loaded.", error);
          });
      })();
    </script>
  </body>
</html>
