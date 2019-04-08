# Install IIS.
dism /online /enable-feature /featurename:IIS-WebServerRole

# Set the home page.
Set-Content `
  -Path "C:\\inetpub\\wwwroot\\Default.htm" `
  -Value "<!DOCTYPE html>
  <html lang="en">

  <head>

      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
      <meta name="description" content="">
      <meta name="author" content="">

      <title>Deloitte Valuations</title>

      <!-- Bootstrap core CSS -->
      <link href="template/bootstrap.min.css" rel="stylesheet">

  </head>

  <body>

      <!-- Navigation -->
      <nav class="navbar navbar-expand-lg navbar-dark bg-dark static-top">
          <div class="container">
              <a class="navbar-brand" href="#">CDO Valuations</a>
              <button class="navbar-toggler" type="button" data-toggle="collapse" data-target="#navbarResponsive" aria-controls="navbarResponsive" aria-expanded="false" aria-label="Toggle navigation">
                  <span class="navbar-toggler-icon"></span>
              </button>
              <div class="collapse navbar-collapse" id="navbarResponsive">
                  <ul class="navbar-nav ml-auto">
                      <li class="nav-item active">
                          <a class="nav-link" href="#">Summary
                              <span class="sr-only">(current)</span>
                          </a>
                      </li>
                      <li class="nav-item">
                          <a class="nav-link" href="https://www.youtube.com/watch?v=7ohbr90cKDc">Breakdown</a>
                      </li>
                      <li class="nav-item">
                          <a class="nav-link" href="portfolio.html">Portfolios</a>
                      </li>
                      <li class="nav-item">
                          <a class="nav-link" href="#">Data</a>
                      </li>
                  </ul>
              </div>
          </div>
      </nav>

      <!-- Page Content -->
      <div class="container">
          <div class="row">
              <div class="col-lg-12 text-center">
                  <h1 class="mt-5">Summary</h1>
                  <p class="lead">An overview of CDO data.</p>

                      <h4>Tranch Value EUR</h4>
                          <img src="vis/Tranch Value EUR.png"/>
                      <p>This figure demonstrates probably very little as I have no idea what any of this is. Nonetheless, observe that the tranch value decreases over time (maybe this means banks are less will to offer high value loans - if that is what these are, I don't know...), very cool.</p>

                      <h4>Tranch Value USD</h4>
                          <img src="vis/Tranch Value USD.png"/>
                      <p>This figure literally shows the exact same thing as the above figure, however this figure is in USD. I should also point out that conceptABS is very shit and that both EUR Equiv. (above) and USD Equiv. (this) are calculated from a Face Value column. Which is great, however what conceptABS can't deal with is super basic formatting differences, so when they have data formatted like [4.5m], rather than converting this into 4,500,000, they just ignore it and leave your EUR Equiv. and USD Equiv. columns blank, so the dataset isn't even complete. Deloitte needs to stop picking shit companies that are stuck in the 90s.</p>

                  <ul class="list-unstyled">
                      <li>Last Updated: 06 Apr 2019</li>
                  </ul>
              </div>
          </div>
      </div>

      <!-- Bootstrap core JavaScript -->
      <script src="template/jquery.min.js"></script>
      <script src="template/bootstrap.bundle.min.js"></script>

  </body>

  </html>"
