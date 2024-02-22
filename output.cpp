
#define SPIE 0
#define SPE 0
#define DORD 0
#define MSTR 0
#define CPOL 0
#define CPHA 0
#define SPR1 0
#define SPR0 0
#define SPI2X 0
#define MSB 0
#define LSB 0

#define HAL_INCLUDES_ADC    1
#define ADC_CHANNEL_REG                     ADC0_MUXPOS
#define ADC_start_conversion()      ADC0_COMMAND |= (1 << ADC_STCONV_bp)
#define ADC_get_status()                    ( !((ADC0_INTFLAGS >> ADC_RESRDY_bp) & 0x01))
#define ADC_get_result()                    ADC0_RES

#define HAL_INCLUDES_DAC    1
#define DAC_enable()                        DAC0_CTRLA |= (1 << DAC_ENABLE_bp)
#define DAC_disable()                       DAC0_CTRLA &= ~(1 << DAC_ENABLE_bp)
#define DAC_output_enable()         DAC0_CTRLA |= (1 << DAC_OUTEN_bp)
#define DAC_output_disable()        DAC0_CTRLA &= ~(1 << DAC_OUTEN_bp)
#define DAC_load(data)                  DAC0_DATA = (data << 6)