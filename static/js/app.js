/**
 * TA RESIDENCE - Utility Monitoring & Billing Dashboard
 * AngularJS 1.8 Client Application
 */
(function () {
  'use strict';

  angular.module('utilityApp', [])

    // ------------------------------------------------------------------------
    // Filters
    // ------------------------------------------------------------------------
    .filter('rupiah', function () {
      return function (input) {
        if (input === null || input === undefined || isNaN(input)) return 'Rp 0';
        return 'Rp ' + Math.round(Number(input)).toLocaleString('id-ID');
      };
    })

    .filter('numberFormat', function () {
      return function (input, decimals) {
        if (input === null || input === undefined || isNaN(input)) return '0';
        var dec = typeof decimals === 'number' ? decimals : 1;
        return Number(input).toLocaleString('id-ID', {
          minimumFractionDigits: 0,
          maximumFractionDigits: dec
        });
      };
    })

    // ------------------------------------------------------------------------
    // Directives
    // ------------------------------------------------------------------------
    // Directive for binding file input to controller model
    .directive('fileModel', ['$parse', function ($parse) {
      return {
        restrict: 'A',
        link: function (scope, element, attrs) {
          var model = $parse(attrs.fileModel);
          var modelSetter = model.assign;

          element.bind('change', function () {
            scope.$apply(function () {
              modelSetter(scope, element[0].files[0] || null);
            });
          });
        }
      };
    }])

    // Directive for Chart.js rendering and updating reactively
    .directive('chartUsage', ['$timeout', function ($timeout) {
      return {
        restrict: 'A',
        scope: {
          chartData: '='
        },
        link: function (scope, element) {
          var chartInstance = null;

          function render(data) {
            if (!data || !data.labels || data.labels.length === 0) {
              if (chartInstance) {
                chartInstance.destroy();
                chartInstance = null;
              }
              return;
            }

            if (chartInstance) {
              chartInstance.destroy();
            }

            var ctx = element[0].getContext('2d');
            chartInstance = new Chart(ctx, {
              type: 'bar',
              data: {
                labels: data.labels,
                datasets: [
                  {
                    label: 'Listrik (kWh)',
                    data: data.elecData,
                    backgroundColor: 'rgba(245, 158, 11, 0.8)',
                    borderRadius: 6,
                    yAxisID: 'y'
                  },
                  {
                    label: 'Air (m³)',
                    data: data.waterData,
                    backgroundColor: 'rgba(6, 182, 212, 0.8)',
                    borderRadius: 6,
                    yAxisID: 'y1'
                  }
                ]
              },
              options: {
                responsive: true,
                maintainAspectRatio: false,
                scales: {
                  y: {
                    type: 'linear',
                    position: 'left',
                    title: { display: true, text: 'Listrik (kWh)' },
                    grid: { color: 'rgba(0, 0, 0, 0.05)' }
                  },
                  y1: {
                    type: 'linear',
                    position: 'right',
                    title: { display: true, text: 'Air (m³)' },
                    grid: { drawOnChartArea: false }
                  }
                }
              }
            });
          }

          scope.$watch('chartData', function (newVal) {
            $timeout(function () {
              render(newVal);
            }, 0);
          }, true);

          scope.$on('$destroy', function () {
            if (chartInstance) {
              chartInstance.destroy();
            }
          });
        }
      };
    }])

    // ------------------------------------------------------------------------
    // Service / Factory
    // ------------------------------------------------------------------------
    .factory('UtilityService', ['$http', function ($http) {
      return {
        getPeriods: function () {
          return $http.get('/api/periods').then(function (res) {
            return res.data;
          });
        },

        getTowers: function () {
          return $http.get('/api/towers').then(function (res) {
            return res.data;
          });
        },

        searchUnits: function (query, tower) {
          return $http.get('/api/units', {
            params: {
              q: query || '',
              tower: tower || ''
            }
          }).then(function (res) {
            return res.data;
          });
        },

        getUsage: function (unit, periods) {
          return $http.get('/api/usage', {
            params: {
              unit: unit,
              periods: periods || []
            }
          }).then(function (res) {
            return res.data;
          });
        },

        exportExcel: function (unit, periods) {
          return $http.get('/api/export-excel', {
            params: {
              unit: unit,
              periods: periods || []
            },
            responseType: 'blob'
          });
        },

        uploadExcel: function (period, file) {
          var formData = new FormData();
          formData.append('period', period);
          formData.append('file', file);

          return $http.post('/api/upload', formData, {
            transformRequest: angular.identity,
            headers: { 'Content-Type': undefined }
          }).then(function (res) {
            return res.data;
          });
        }
      };
    }])

    // ------------------------------------------------------------------------
    // Controller
    // ------------------------------------------------------------------------
    .controller('UtilityController', [
      '$scope',
      '$timeout',
      '$window',
      'UtilityService',
      function ($scope, $timeout, $window, UtilityService) {
        var vm = this;

        // State
        vm.availablePeriods = [];
        vm.selectedPeriods = [];
        vm.selectedTower = '';
        vm.searchQuery = '';
        vm.searchResults = [];
        vm.showDropdown = false;
        vm.isSearching = false;

        vm.currentUnit = null; // { unit, owner_name, tower, kva }
        vm.records = [];

        vm.metrics = {
          elecUsage: 0,
          elecCost: 0,
          waterUsage: 0,
          waterCost: 0,
          grandTotal: 0,
          periodCount: 0
        };

        vm.chartData = {
          labels: [],
          elecData: [],
          waterData: []
        };

        vm.isExporting = false;

        // Upload modal state
        vm.uploadModalOpen = false;
        vm.isUploading = false;
        vm.uploadPeriod = '';
        vm.uploadFile = null;
        vm.uploadStatus = {
          visible: false,
          type: 'info', // 'info' | 'success' | 'error'
          message: ''
        };

        var searchTimeout = null;

        // --------------------------------------------------------------------
        // Initialization
        // --------------------------------------------------------------------
        vm.init = function () {
          vm.fetchPeriods();
          vm.refreshIcons();
        };

        vm.refreshIcons = function () {
          $timeout(function () {
            if ($window.lucide && typeof $window.lucide.createIcons === 'function') {
              $window.lucide.createIcons();
            }
          }, 50);
        };

        vm.fetchPeriods = function () {
          UtilityService.getPeriods()
            .then(function (periods) {
              vm.availablePeriods = periods;
              // Default select all available periods
              vm.selectedPeriods = angular.copy(periods);
              if (vm.currentUnit) {
                vm.loadUsage();
              }
            })
            .catch(function (err) {
              console.error('Gagal mengambil daftar periode:', err);
            });
        };

        // --------------------------------------------------------------------
        // Filter & Selection
        // --------------------------------------------------------------------
        vm.isPeriodSelected = function (period) {
          return vm.selectedPeriods.indexOf(period) > -1;
        };

        vm.togglePeriod = function (period) {
          var index = vm.selectedPeriods.indexOf(period);
          if (index > -1) {
            vm.selectedPeriods.splice(index, 1);
          } else {
            vm.selectedPeriods.push(period);
          }

          if (vm.currentUnit) {
            vm.loadUsage();
          }
        };

        vm.onSearchInput = function () {
          if (searchTimeout) {
            $timeout.cancel(searchTimeout);
          }

          if (!vm.searchQuery && !vm.selectedTower) {
            vm.searchResults = [];
            vm.showDropdown = false;
            return;
          }

          searchTimeout = $timeout(function () {
            vm.performSearch();
          }, 250);
        };

        vm.onTowerChange = function () {
          vm.performSearch();
        };

        vm.performSearch = function () {
          vm.isSearching = true;
          UtilityService.searchUnits(vm.searchQuery, vm.selectedTower)
            .then(function (units) {
              vm.searchResults = units;
              vm.showDropdown = true;
            })
            .catch(function (err) {
              console.error('Gagal mencari unit:', err);
              vm.searchResults = [];
            })
            .finally(function () {
              vm.isSearching = false;
              vm.refreshIcons();
            });
        };

        vm.selectUnit = function (unitObj) {
          vm.currentUnit = unitObj;
          vm.searchQuery = unitObj.unit + ' - ' + (unitObj.owner_name || '');
          vm.showDropdown = false;
          vm.searchResults = [];

          vm.loadUsage();
          vm.refreshIcons();
        };

        vm.closeDropdown = function () {
          $timeout(function () {
            vm.showDropdown = false;
          }, 200);
        };

        // --------------------------------------------------------------------
        // Usage Data Loading & Calculation
        // --------------------------------------------------------------------
        vm.loadUsage = function () {
          if (!vm.currentUnit || !vm.currentUnit.unit) return;

          UtilityService.getUsage(vm.currentUnit.unit, vm.selectedPeriods)
            .then(function (records) {
              vm.records = records || [];
              vm.calculateMetrics(vm.records);
              vm.refreshIcons();
            })
            .catch(function (err) {
              console.error('Gagal mengambil data pemakaian:', err);
              vm.records = [];
              vm.calculateMetrics([]);
            });
        };

        vm.calculateMetrics = function (records) {
          var totalElecUsage = 0;
          var totalElecCost = 0;
          var totalWaterUsage = 0;
          var totalWaterCost = 0;
          var totalGrand = 0;

          var labels = [];
          var elecData = [];
          var waterData = [];

          angular.forEach(records, function (r) {
            totalElecUsage += (r.elec_usage || 0);
            totalElecCost += (r.elec_total || 0);
            totalWaterUsage += (r.water_usage || 0);
            totalWaterCost += (r.water_total || 0);
            totalGrand += (r.grand_total || 0);

            labels.push(r.period);
            elecData.push(r.elec_usage || 0);
            waterData.push(r.water_usage || 0);
          });

          vm.metrics.elecUsage = totalElecUsage;
          vm.metrics.elecCost = totalElecCost;
          vm.metrics.waterUsage = totalWaterUsage;
          vm.metrics.waterCost = totalWaterCost;
          vm.metrics.grandTotal = totalGrand;
          vm.metrics.periodCount = records.length;

          vm.chartData = {
            labels: labels,
            elecData: elecData,
            waterData: waterData
          };
        };

        // --------------------------------------------------------------------
        // Export Excel
        // --------------------------------------------------------------------
        vm.exportToExcel = function () {
          if (!vm.currentUnit || !vm.currentUnit.unit) {
            $window.alert('Silakan pilih unit terlebih dahulu sebelum export Excel.');
            return;
          }

          if (vm.selectedPeriods.length === 0) {
            $window.alert('Silakan pilih minimal 1 periode bulan untuk diekspor.');
            return;
          }

          vm.isExporting = true;

          UtilityService.exportExcel(vm.currentUnit.unit, vm.selectedPeriods)
            .then(function (response) {
              var blob = new Blob([response.data], {
                type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
              });

              var filename = 'Laporan_Utilitas_' + vm.currentUnit.unit + '.xlsx';
              var disposition = response.headers('content-disposition');
              if (disposition) {
                var match = disposition.match(/filename\*?=(?:UTF-8'')?["']?([^"';\n]+)["']?/i);
                if (match && match[1]) {
                  filename = decodeURIComponent(match[1]);
                }
              }

              var downloadUrl = $window.URL.createObjectURL(blob);
              var link = document.createElement('a');
              link.href = downloadUrl;
              link.download = filename;
              document.body.appendChild(link);
              link.click();
              link.remove();

              $timeout(function () {
                $window.URL.revokeObjectURL(downloadUrl);
              }, 1000);
            })
            .catch(function (err) {
              console.error('Gagal export excel:', err);
              $window.alert('Terjadi kesalahan saat mengekspor file Excel.');
            })
            .finally(function () {
              vm.isExporting = false;
              vm.refreshIcons();
            });
        };

        // --------------------------------------------------------------------
        // Print Slip
        // --------------------------------------------------------------------
        vm.printSlip = function () {
          $window.print();
        };

        // --------------------------------------------------------------------
        // Upload Excel Modal
        // --------------------------------------------------------------------
        vm.openUploadModal = function () {
          vm.uploadModalOpen = true;
          vm.uploadPeriod = '';
          vm.uploadFile = null;
          vm.uploadStatus = { visible: false, type: 'info', message: '' };
          vm.refreshIcons();
        };

        vm.closeUploadModal = function () {
          vm.uploadModalOpen = false;
          vm.isUploading = false;
          vm.uploadFile = null;
          vm.uploadStatus = { visible: false, type: 'info', message: '' };
        };

        vm.submitUpload = function () {
          if (!vm.uploadPeriod || !vm.uploadFile) {
            vm.uploadStatus = {
              visible: true,
              type: 'error',
              message: 'Mohon isi periode (YYYY-MM) dan pilih file .xlsx!'
            };
            return;
          }

          vm.isUploading = true;
          vm.uploadStatus = {
            visible: true,
            type: 'info',
            message: 'Sedang memproses dan mengimpor data ke database...'
          };

          UtilityService.uploadExcel(vm.uploadPeriod, vm.uploadFile)
            .then(function (result) {
              vm.uploadStatus = {
                visible: true,
                type: 'success',
                message: result.message || 'Berhasil mengimpor data!'
              };

              $timeout(function () {
                vm.closeUploadModal();
                vm.fetchPeriods();
                if (vm.currentUnit) {
                  vm.loadUsage();
                }
              }, 1500);
            })
            .catch(function (err) {
              var errMsg = 'Terjadi kesalahan saat upload.';
              if (err.data && err.data.error) {
                errMsg = err.data.error;
              }
              vm.uploadStatus = {
                visible: true,
                type: 'error',
                message: errMsg
              };
            })
            .finally(function () {
              vm.isUploading = false;
              vm.refreshIcons();
            });
        };

        // Start initialization
        vm.init();
      }
    ]);
})();
