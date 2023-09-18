
_validation = {
    initializeValidation: function ()
    {
      var form = $("#needs-validation");
      var submit = form.find(".submit");
      form.on('submit', function(e) {

          if (!this.checkValidity()) {
            e.preventDefault();
            e.stopPropagation();
          }

          $(this).addClass('was-validated');
        });      
    }
  };
_transitions = {

    initializeDirectionalFadeIn: function ()
    {
        var bottomOfScreen = $(window).scrollTop() + $(window).height();
        var topOfScreen = $(window).scrollTop();        
        var element = $('.fade-in');
        element.each(function ()
        {
            var element = $(this);
            if (!element.hasClass("active"))
            {                
                var top = element.offset().top + 40;
                if (bottomOfScreen > top)
                {
                    element.addClass("active");
                }  
            }
        });
    }
  };

$(window).scroll(function () {
   _transitions.initializeDirectionalFadeIn();
});

$(document).ready(function (){
    _validation.initializeValidation();
});